

Imports ProjectBoardDefinitions
'Imports DBAccLayer
Imports ProjectboardReports
Imports Microsoft.Office.Interop
Imports Microsoft.Office.Interop.Excel
'Imports System.ComponentModel
'Imports System.Windows
Imports System.Windows.Forms
Imports System.Security.Principal
Imports System.Text.RegularExpressions

Public Module agm3

    ''' <summary>
    ''' überprüft, ob die Voraussetzungen für das Einlesen der InternenAnwesenheitslisten. 
    ''' </summary>
    ''' <param name="configFile"></param>
    ''' <param name="kapaFile"></param>
    ''' <param name="kapaConfigs"></param>
    ''' <param name="lastrow"></param>
    ''' <returns></returns>
    Public Function checkCapaImportConfig(ByVal configFile As String,
                                      ByRef kapaFile As String,
                                      ByRef kapaConfigs As SortedList(Of String, clsConfigKapaImport),
                                      ByRef lastrow As Integer, ByRef oPCollection As Collection) As Boolean

        Dim outputline As String = ""
        Dim configLine As New clsConfigKapaImport
        Dim currentDirectoryName As String = requirementsOrdner
        Dim configWB As Microsoft.Office.Interop.Excel.Workbook = Nothing
        Dim currentWS As Microsoft.Office.Interop.Excel.Worksheet = Nothing
        Dim searcharea As Microsoft.Office.Interop.Excel.Range = Nothing
        Dim anzOld_oPCollection As Integer = oPCollection.Count
        'Dim found As Boolean
        'Dim i As Integer

        ''
        '' Config-file wird geöffnet
        ' Filename ggf. mit Directory erweitern
        configFile = My.Computer.FileSystem.CombinePath(currentDirectoryName, configFile)

        ' öffnen des Files 
        If My.Computer.FileSystem.FileExists(configFile) Then

            Try
                configWB = appInstance.Workbooks.Open(configFile)

                Try

                    If appInstance.Worksheets.Count > 0 Then

                        'currentWS = CType(appInstance.Worksheets(1), Global.Microsoft.Office.Interop.Excel.Worksheet)
                        currentWS = CType(configWB.Worksheets("VISBO Config"), Global.Microsoft.Office.Interop.Excel.Worksheet)


                        Dim titleCol As Integer,
                            IdentCol As Integer,
                            InputFileCol As Integer,
                            TypCol As Integer,
                            DatenCol As Integer,
                            TabUCol As Integer, TabNCol As Integer,
                            SUCol As Integer, SNCol As Integer,
                            ZUCol As Integer, ZNCol As Integer,
                            ObjCol As Integer,
                            InhaltCol As Integer

                        ' ImportTyp aus configfile lesen, wenn nicht vorhanden, wird es übergangen
                        configLine = New clsConfigKapaImport
                        Dim titelzeile As Integer = 4  ' ursprüngliche Zeile der Titel war 4 aber nach Import-Änderung für Instart ist in Zeile 4 der Importtyp verankert
                        configLine.Titel = CStr(currentWS.Cells(titelzeile, 1).value)
                        configLine.content = CStr(currentWS.Cells(titelzeile, 2).value)

                        If Not IsNothing(configLine.Titel) Then
                            kapaConfigs.Add(configLine.Titel, configLine)
                        End If


                        searcharea = currentWS.Rows(5)          ' Zeile 5 enthält die verschieden Configurationselemente

                        titleCol = searcharea.Find("Titel").Column
                        IdentCol = searcharea.Find("Identifier").Column
                        InputFileCol = searcharea.Find("InputFile").Column
                        TypCol = searcharea.Find("Typ").Column
                        DatenCol = searcharea.Find("Datenbereich").Column
                        TabUCol = searcharea.Find("Tabellen-Name").Column
                        TabNCol = searcharea.Find("Tabellen-Nummer").Column
                        SUCol = searcharea.Find("Spaltenüberschrift").Column
                        SNCol = searcharea.Find("Spalten-Nummer").Column
                        ZUCol = searcharea.Find("Zeilenbeschriftung").Column
                        ZNCol = searcharea.Find("Zeilen-Nummer").Column
                        ObjCol = searcharea.Find("Objekt-Typ").Column
                        InhaltCol = searcharea.Find("Inhalt").Column

                        Dim ok As Boolean = (titleCol + IdentCol + TypCol + DatenCol + SUCol + SNCol + ZUCol + ZNCol + ObjCol + InhaltCol > 13)


                        If ok Then
                            With currentWS

                                lastrow = .Cells(.Rows.Count, titleCol).end(Microsoft.Office.Interop.Excel.XlDirection.xlUp).row

                                For i = 6 To lastrow

                                    configLine = New clsConfigKapaImport

                                    Dim Titel As String = CStr(.Cells(i, titleCol).value)

                                    Select Case Titel
                                        Case "Kapa-Datei"
                                            configLine.Titel = CStr(.Cells(i, titleCol).value)
                                            configLine.capacityFile = CStr(.Cells(i, InputFileCol).value)
                                            kapaFile = configLine.capacityFile

                                        Case "month"
                                            configLine.Titel = CStr(.Cells(i, titleCol).value)
                                            configLine.Identifier = CStr(.Cells(i, IdentCol).value)
                                            configLine.Inputfile = CStr(.Cells(i, InputFileCol).value)
                                            configLine.Typ = CStr(.Cells(i, TypCol).value)
                                            configLine.cellrange = (CStr(.Cells(i, DatenCol).value) = "Range")
                                            configLine.tabNr = CInt(.Cells(i, TabNCol).value)
                                            configLine.tabName = CStr(.Cells(i, TabUCol).value)
                                            configLine.column = CInt(.Cells(i, SNCol).value)
                                            configLine.columnDescript = CStr(.Cells(i, SUCol).value)
                                            configLine.row = CInt(.Cells(i, ZNCol).value)
                                            configLine.rowDescript = CStr(.Cells(i, ZUCol).value)
                                            configLine.regex = CStr(.Cells(i, ObjCol).value)
                                            configLine.content = CStr(.Cells(i, InhaltCol).value)

                                        Case "year"
                                            configLine.Titel = CStr(.Cells(i, titleCol).value)
                                            configLine.Identifier = CStr(.Cells(i, IdentCol).value)
                                            configLine.Inputfile = CStr(.Cells(i, InputFileCol).value)
                                            configLine.Typ = CStr(.Cells(i, TypCol).value)
                                            configLine.cellrange = (CStr(.Cells(i, DatenCol).value) = "Range")
                                            configLine.tabNr = CInt(.Cells(i, TabNCol).value)
                                            configLine.tabName = CStr(.Cells(i, TabUCol).value)
                                            configLine.column = CInt(.Cells(i, SNCol).value)
                                            configLine.columnDescript = CStr(.Cells(i, SUCol).value)
                                            configLine.row = CInt(.Cells(i, ZNCol).value)
                                            configLine.rowDescript = CStr(.Cells(i, ZUCol).value)
                                            configLine.regex = CStr(.Cells(i, ObjCol).value)
                                            configLine.content = CStr(.Cells(i, InhaltCol).value)

                                        Case "role"
                                            configLine.Titel = CStr(.Cells(i, titleCol).value)
                                            configLine.Identifier = CStr(.Cells(i, IdentCol).value)
                                            configLine.Inputfile = CStr(.Cells(i, InputFileCol).value)
                                            configLine.Typ = CStr(.Cells(i, TypCol).value)
                                            configLine.cellrange = (CStr(.Cells(i, DatenCol).value) = "Range")
                                            configLine.tabNr = CInt(.Cells(i, TabNCol).value)
                                            configLine.tabName = CStr(.Cells(i, TabUCol).value)
                                            configLine.column = CInt(.Cells(i, SNCol).value)
                                            configLine.columnDescript = CStr(.Cells(i, SUCol).value)
                                            configLine.row = CInt(.Cells(i, ZNCol).value)
                                            configLine.rowDescript = CStr(.Cells(i, ZUCol).value)
                                            configLine.regex = CStr(.Cells(i, ObjCol).value)
                                            configLine.content = CStr(.Cells(i, InhaltCol).value)


                                        Case "valueStart"
                                            configLine.Titel = CStr(.Cells(i, titleCol).value)
                                            configLine.Identifier = CStr(.Cells(i, IdentCol).value)
                                            configLine.Inputfile = CStr(.Cells(i, InputFileCol).value)
                                            configLine.Typ = CStr(.Cells(i, TypCol).value)
                                            configLine.cellrange = (CStr(.Cells(i, DatenCol).value) = "Range")
                                            configLine.tabNr = CInt(.Cells(i, TabNCol).value)
                                            configLine.tabName = CStr(.Cells(i, TabUCol).value)
                                            configLine.column = CInt(.Cells(i, SNCol).value)
                                            configLine.columnDescript = CStr(.Cells(i, SUCol).value)
                                            configLine.row = CInt(.Cells(i, ZNCol).value)
                                            configLine.rowDescript = CStr(.Cells(i, ZUCol).value)
                                            configLine.regex = CStr(.Cells(i, ObjCol).value)
                                            configLine.content = CStr(.Cells(i, InhaltCol).value)

                                        Case "valueLength"
                                            configLine.Titel = CStr(.Cells(i, titleCol).value)
                                            configLine.Identifier = CStr(.Cells(i, IdentCol).value)
                                            configLine.Inputfile = CStr(.Cells(i, InputFileCol).value)
                                            configLine.Typ = CStr(.Cells(i, TypCol).value)
                                            configLine.cellrange = (CStr(.Cells(i, DatenCol).value) = "Range")
                                            configLine.tabNr = CInt(.Cells(i, TabNCol).value)
                                            configLine.tabName = CStr(.Cells(i, TabUCol).value)
                                            configLine.column = CInt(.Cells(i, SNCol).value)
                                            configLine.columnDescript = CStr(.Cells(i, SUCol).value)
                                            configLine.row = CInt(.Cells(i, ZNCol).value)
                                            configLine.rowDescript = CStr(.Cells(i, ZUCol).value)
                                            configLine.regex = CStr(.Cells(i, ObjCol).value)
                                            configLine.content = CStr(.Cells(i, InhaltCol).value)

                                        Case "valueSign"
                                            configLine.Titel = CStr(.Cells(i, titleCol).value)
                                            configLine.Identifier = CStr(.Cells(i, IdentCol).value)
                                            configLine.Inputfile = CStr(.Cells(i, InputFileCol).value)
                                            configLine.Typ = CStr(.Cells(i, TypCol).value)
                                            configLine.cellrange = (CStr(.Cells(i, DatenCol).value) = "Range")
                                            configLine.tabNr = CInt(.Cells(i, TabNCol).value)
                                            configLine.tabName = CStr(.Cells(i, TabUCol).value)
                                            configLine.column = CInt(.Cells(i, SNCol).value)
                                            configLine.columnDescript = CStr(.Cells(i, SUCol).value)
                                            configLine.row = CInt(.Cells(i, ZNCol).value)
                                            configLine.rowDescript = CStr(.Cells(i, ZUCol).value)
                                            configLine.regex = CStr(.Cells(i, ObjCol).value)
                                            configLine.content = CStr(.Cells(i, InhaltCol).value)


                                        Case "LastLine"
                                            configLine.Titel = CStr(.Cells(i, titleCol).value)
                                            configLine.Identifier = CStr(.Cells(i, IdentCol).value)
                                            configLine.Inputfile = CStr(.Cells(i, InputFileCol).value)
                                            configLine.Typ = CStr(.Cells(i, TypCol).value)
                                            configLine.cellrange = (CStr(.Cells(i, DatenCol).value) = "Range")
                                            configLine.tabNr = CInt(.Cells(i, TabNCol).value)
                                            configLine.tabName = CStr(.Cells(i, TabUCol).value)
                                            configLine.column = CInt(.Cells(i, SNCol).value)
                                            configLine.columnDescript = CStr(.Cells(i, SUCol).value)
                                            configLine.row = CInt(.Cells(i, ZNCol).value)
                                            configLine.rowDescript = CStr(.Cells(i, ZUCol).value)
                                            configLine.regex = CStr(.Cells(i, ObjCol).value)
                                            configLine.content = CStr(.Cells(i, InhaltCol).value)

                                        Case Else
                                            configLine.Titel = CStr(.Cells(i, titleCol).value)
                                            configLine.Identifier = CStr(.Cells(i, IdentCol).value)
                                            configLine.Inputfile = CStr(.Cells(i, InputFileCol).value)
                                            configLine.Typ = CStr(.Cells(i, TypCol).value)
                                            configLine.cellrange = (CStr(.Cells(i, DatenCol).value) = "Range")
                                            configLine.tabNr = CInt(.Cells(i, TabNCol).value)
                                            configLine.tabName = CStr(.Cells(i, TabUCol).value)
                                            configLine.column = CInt(.Cells(i, SNCol).value)
                                            configLine.columnDescript = CStr(.Cells(i, SUCol).value)
                                            configLine.row = CInt(.Cells(i, ZNCol).value)
                                            configLine.rowDescript = CStr(.Cells(i, ZUCol).value)
                                            configLine.regex = CStr(.Cells(i, ObjCol).value)
                                            configLine.content = CStr(.Cells(i, InhaltCol).value)

                                    End Select

                                    If kapaConfigs.ContainsKey(configLine.Titel) Then
                                        kapaConfigs.Remove(configLine.Titel)
                                    End If

                                    kapaConfigs.Add(configLine.Titel, configLine)

                                Next

                            End With
                        Else
                            outputline = "Die Konfigurationsdatei stimmt nicht mit der erwarteten Struktur überein!"
                            If awinSettings.englishLanguage Then
                                outputline = "Configuration file does not have expected structure! please contact your sys-admin or VISBO"
                            End If
                            oPCollection.Add(outputline)
                        End If
                    End If

                Catch ex As Exception
                    If awinSettings.englishLanguage Then
                        outputline = "The configurationfile " & configFile & " has no Sheet with name VISBO Config" & vbCrLf & " ... no import!"
                    Else
                        outputline = "Die Konfigurationsdatei " & configFile & " enthält kein Registerblatt VISBO Config" &
                                    vbCrLf & " es fand kein Import statt "
                    End If
                    oPCollection.Add(outputline)
                End Try

                ' configCapaImport - Konfigurationsfile schließen
                configWB.Close(SaveChanges:=False)

            Catch ex As Exception
                outputline = "Die Konfigurationsdatei konnte nicht geöffnet werden - " & configFile
                If awinSettings.englishLanguage Then
                    outputline = "Config File could not be opened - please contact your sys-admin or VISBO"
                End If
                oPCollection.Add(outputline)
                'Call MsgBox(outputline)
            End Try
        Else
            ' soll nur Info im Logbuch sein
            outputline = "Keine Konfigurationsdatei für Import Capacities vorhanden! - " & configFile
            If awinSettings.englishLanguage Then
                outputline = "There is no such config file: " & configFile
            End If
            Call logger(ptErrLevel.logWarning, outputline, "", -1)
        End If

        checkCapaImportConfig = (kapaConfigs.Count > 0) And (anzOld_oPCollection - oPCollection.Count = 0)

    End Function


    ''' <summary>
    ''' überprüft, ob die Voraussetzungen für das Einlesen der Projekte. 
    ''' </summary>
    ''' <param name="configFile"></param>
    ''' <param name="ProjectsFile"></param>
    ''' <param name="ProjectsConfigs"></param>
    ''' <param name="lastrow"></param>
    ''' <returns></returns>
    Public Function checkProjectImportConfig(ByVal configFile As String,
                                      ByRef ProjectsFile As String,
                                      ByRef ProjectsConfigs As SortedList(Of String, clsConfigProjectsImport),
                                      ByRef lastrow As Integer,
                                      ByRef outputCollection As Collection) As Boolean


        Dim configLine As New clsConfigProjectsImport
        Dim currentDirectoryName As String = requirementsOrdner
        Dim configWB As Microsoft.Office.Interop.Excel.Workbook = Nothing
        Dim currentWS As Microsoft.Office.Interop.Excel.Worksheet = Nothing
        Dim searcharea As Microsoft.Office.Interop.Excel.Range = Nothing
        Dim outputLine As String

        ''
        '' Config-file wird geöffnet
        ' Filename ggf. mit Directory erweitern
        configFile = My.Computer.FileSystem.CombinePath(currentDirectoryName, configFile)

        ' öffnen des Files 
        If My.Computer.FileSystem.FileExists(configFile) Then

            Try
                configWB = appInstance.Workbooks.Open(configFile)

                Try

                    If appInstance.Worksheets.Count > 0 Then

                        'currentWS = CType(appInstance.Worksheets(1), Global.Microsoft.Office.Interop.Excel.Worksheet)
                        currentWS = CType(configWB.Worksheets("VISBO Config"), Global.Microsoft.Office.Interop.Excel.Worksheet)

                        Dim titleCol As Integer,
                            IdentCol As Integer,
                            InputFileCol As Integer,
                            TypCol As Integer,
                            DatenCol As Integer,
                            TabUCol As Integer, TabNCol As Integer,
                            SUCol As Integer, SNCol As Integer,
                            ZUCol As Integer, ZNCol As Integer,
                            ObjCol As Integer,
                            InhaltCol As Integer

                        searcharea = currentWS.Rows(5)          ' Zeile 5 enthält die verschieden Configurationselemente

                        titleCol = searcharea.Find("Titel").Column
                        IdentCol = searcharea.Find("Identifier").Column
                        InputFileCol = searcharea.Find("InputFile").Column
                        TypCol = searcharea.Find("Typ").Column
                        DatenCol = searcharea.Find("Datenbereich").Column
                        TabUCol = searcharea.Find("Tabellen-Name").Column
                        TabNCol = searcharea.Find("Tabellen-Nummer").Column
                        SUCol = searcharea.Find("Spaltenüberschrift").Column
                        SNCol = searcharea.Find("Spalten-Nummer").Column
                        ZUCol = searcharea.Find("Zeilenbeschriftung").Column
                        ZNCol = searcharea.Find("Zeilen-Nummer").Column
                        ObjCol = searcharea.Find("Objekt-Typ").Column
                        InhaltCol = searcharea.Find("Inhalt").Column

                        Dim ok As Boolean = (titleCol + IdentCol + TypCol + DatenCol + SUCol + SNCol + ZUCol + ZNCol + ObjCol + InhaltCol > 13)

                        If ok Then
                            With currentWS
                                lastrow = .Cells(.Rows.Count, titleCol).end(Microsoft.Office.Interop.Excel.XlDirection.xlUp).row

                                For i = 6 To lastrow

                                    configLine = New clsConfigProjectsImport

                                    Dim Titel As String = CStr(.Cells(i, titleCol).value)

                                    Select Case Titel
                                        Case "DateiName"
                                            configLine.Titel = CStr(.Cells(i, titleCol).value)
                                            configLine.ProjectsFile = CStr(.Cells(i, InputFileCol).value)
                                            ProjectsFile = configLine.ProjectsFile


                                        Case Else
                                            configLine.Titel = CStr(.Cells(i, titleCol).value)
                                            configLine.Identifier = CStr(.Cells(i, IdentCol).value)
                                            configLine.Inputfile = CStr(.Cells(i, InputFileCol).value)
                                            configLine.Typ = CStr(.Cells(i, TypCol).value)
                                            configLine.cellrange = (CStr(.Cells(i, DatenCol).value) = "Range")
                                            configLine.sheet = CInt(.Cells(i, TabNCol).value)
                                            configLine.sheetDescript = CStr(.Cells(i, TabUCol).value)
                                            If configLine.cellrange Then
                                                Dim colrange As String = CStr(.Cells(i, SNCol).value)
                                                Dim hstr() As String = Split(colrange, ":")
                                                If hstr.Length = 2 Then
                                                    configLine.column.von = CInt(hstr(0))
                                                    configLine.column.bis = CInt(hstr(1))
                                                ElseIf hstr.Length = 1 Then
                                                    configLine.row.von = CInt(.Cells(i, SNCol).value)
                                                    configLine.row.bis = CInt(.Cells(i, SNCol).value)
                                                Else
                                                    outputLine = configLine.Titel & " : Angabe ist kein Range"
                                                End If
                                            Else
                                                configLine.column.von = CInt(.Cells(i, SNCol).value)
                                                configLine.column.bis = CInt(.Cells(i, SNCol).value)
                                            End If
                                            configLine.columnDescript = CStr(.Cells(i, SUCol).value)

                                            If configLine.cellrange Then
                                                Dim colrange As String = CStr(.Cells(i, ZNCol).value)
                                                Dim hstr() As String = Split(colrange, ":")
                                                If hstr.Length = 2 Then
                                                    configLine.row.von = CInt(hstr(0))
                                                    configLine.row.bis = CInt(hstr(1))
                                                ElseIf hstr.Length = 1 Then
                                                    configLine.row.von = CInt(.Cells(i, ZNCol).value)
                                                    configLine.row.bis = CInt(.Cells(i, ZNCol).value)
                                                Else
                                                    outputLine = configLine.Titel & " : Angabe ist kein Range"
                                                End If
                                            Else
                                                configLine.row.von = CInt(.Cells(i, ZNCol).value)
                                                configLine.row.bis = CInt(.Cells(i, ZNCol).value)
                                            End If
                                            configLine.rowDescript = CStr(.Cells(i, ZUCol).value)
                                            configLine.objType = CStr(.Cells(i, ObjCol).value)
                                            configLine.content = CStr(.Cells(i, InhaltCol).value)
                                    End Select

                                    If ProjectsConfigs.ContainsKey(configLine.Titel) Then
                                        ProjectsConfigs.Remove(configLine.Titel)
                                    End If

                                    ProjectsConfigs.Add(configLine.Titel, configLine)

                                Next

                            End With
                        Else
                            If awinSettings.englishLanguage Then
                                outputLine = "The structure of the configFile doesn't match!  -  " & configFile
                            Else
                                outputLine = "Der Aufbau der Konfigurationsdatei ist nicht passend  -  " & configFile
                            End If
                            outputCollection.Add(outputLine)
                        End If

                    End If

                Catch ex As Exception
                    If awinSettings.englishLanguage Then
                        outputLine = "The configurationfile " & configFile & " has no Sheet with name VISBO Config" & vbCrLf & " ... no import!"
                    Else
                        outputLine = "Die Konfigurationsdatei " & configFile & " enthält kein Registerblatt VISBO Config" &
                                    vbCrLf & " es fand kein Import statt "
                    End If
                    outputCollection.Add(outputLine)
                End Try

                ' configCapaImport - Konfigurationsfile schließen
                configWB.Close(SaveChanges:=False)

            Catch ex As Exception
                If awinSettings.englishLanguage Then
                    Call MsgBox("The configuration-file " & configFile & "  To import the projects couldn't be opened.")
                    outputLine = "The configurationfile " & configFile & "  to import the projects couldn't be opened."
                Else
                    Call MsgBox("Das Öffnen der Konfigurationsdatei " & configFile & " war nicht erfolgreich." &
                                vbCrLf & " Die Projekte können somit nicht importiert werden")
                    outputLine = "Das Öffnen der Konfigurationsdatei " & configFile & " war nicht erfolgreich." &
                                vbCrLf & " Die Projekte können somit nicht importiert werden"
                End If
                outputCollection.Add(outputLine)
            End Try
        Else
            If awinSettings.englishLanguage Then
                outputLine = "The configuration-file doen't exist!  -  " & configFile
            Else
                outputLine = "Die Konfigurationsdatei existiert nicht!  -  " & configFile
            End If
            outputCollection.Add(outputLine)
        End If

        checkProjectImportConfig = (ProjectsConfigs.Count > 0)

    End Function

    ''' <summary>
    ''' überprüft, ob die Voraussetzungen für das Einlesen der Projekte. 
    ''' </summary>
    ''' <param name="configFile"></param>
    ''' <param name="ActualDataFile"></param>
    ''' <param name="ActualDataConfigs"></param>
    ''' <param name="lastrow"></param>
    ''' <returns></returns>
    Public Function checkActualDataImportConfig(ByVal configFile As String,
                                      ByRef ActualDataFile As String,
                                      ByRef ActualDataConfigs As SortedList(Of String, clsConfigActualDataImport),
                                      ByRef lastrow As Integer,
                                      ByRef outputCollection As Collection) As Boolean

        Dim configLine As New clsConfigActualDataImport
        Dim currentDirectoryName As String = requirementsOrdner
        Dim configWB As Microsoft.Office.Interop.Excel.Workbook = Nothing
        Dim currentWS As Microsoft.Office.Interop.Excel.Worksheet = Nothing
        Dim searcharea As Microsoft.Office.Interop.Excel.Range = Nothing
        Dim outputLine As String

        ''
        '' Config-file wird geöffnet
        ' Filename ggf. mit Directory erweitern
        configFile = My.Computer.FileSystem.CombinePath(currentDirectoryName, configFile)

        ' öffnen des Files 
        If My.Computer.FileSystem.FileExists(configFile) Then

            Try

                configWB = appInstance.Workbooks.Open(configFile)

                Try
                    If appInstance.Worksheets.Count > 0 Then

                        'currentWS = CType(appInstance.Worksheets(1), Global.Microsoft.Office.Interop.Excel.Worksheet)
                        currentWS = CType(configWB.Worksheets("VISBO Config"), Global.Microsoft.Office.Interop.Excel.Worksheet)

                        Dim titleCol As Integer,
                            IdentCol As Integer,
                            InputFileCol As Integer,
                            TypCol As Integer,
                            DatenCol As Integer,
                            TabUCol As Integer, TabNCol As Integer,
                            SUCol As Integer, SNCol As Integer,
                            ZUCol As Integer, ZNCol As Integer,
                            ObjCol As Integer,
                            InhaltCol As Integer

                        searcharea = currentWS.Rows(5)          ' Zeile 5 enthält die verschieden Configurationselemente

                        titleCol = searcharea.Find("Titel").Column
                        IdentCol = searcharea.Find("Identifier").Column
                        InputFileCol = searcharea.Find("InputFile").Column
                        TypCol = searcharea.Find("Typ").Column
                        DatenCol = searcharea.Find("Datenbereich").Column
                        TabUCol = searcharea.Find("Tabellen-Name").Column
                        TabNCol = searcharea.Find("Tabellen-Nummer").Column
                        SUCol = searcharea.Find("Spaltenüberschrift").Column
                        SNCol = searcharea.Find("Spalten-Nummer").Column
                        ZUCol = searcharea.Find("Zeilenbeschriftung").Column
                        ZNCol = searcharea.Find("Zeilen-Nummer").Column
                        ObjCol = searcharea.Find("Objekt-Typ").Column
                        InhaltCol = searcharea.Find("Inhalt").Column

                        Dim ok As Boolean = (titleCol + IdentCol + TypCol + DatenCol + SUCol + SNCol + ZUCol + ZNCol + ObjCol + InhaltCol > 13)

                        If ok Then
                            With currentWS
                                lastrow = .Cells(.Rows.Count, titleCol).end(Microsoft.Office.Interop.Excel.XlDirection.xlUp).row

                                For i = 6 To lastrow

                                    configLine = New clsConfigActualDataImport

                                    Dim Titel As String = CStr(.Cells(i, titleCol).value)

                                    Select Case Titel
                                        Case "DateiName"
                                            configLine.Titel = CStr(.Cells(i, titleCol).value)
                                            configLine.Inputfile = CStr(.Cells(i, InputFileCol).value)
                                            ActualDataFile = configLine.Inputfile



                                        Case Else
                                            configLine.Titel = CStr(.Cells(i, titleCol).value)
                                            configLine.Identifier = CStr(.Cells(i, IdentCol).value)
                                            configLine.Inputfile = CStr(.Cells(i, InputFileCol).value)
                                            configLine.Typ = CStr(.Cells(i, TypCol).value)
                                            configLine.cellrange = (CStr(.Cells(i, DatenCol).value) = "Range")
                                            'configLine.sheet = CInt(.Cells(i, TabNCol).value)
                                            Dim tabrange As String = CStr(.Cells(i, TabNCol).value)
                                            Dim hstr() As String = Split(tabrange, ":")
                                            If hstr.Length = 2 Then
                                                configLine.sheet.von = CInt(hstr(0))
                                                configLine.sheet.bis = CInt(hstr(1))
                                            ElseIf hstr.Length = 1 Then
                                                configLine.sheet.von = CInt(.Cells(i, SNCol).value)
                                                configLine.sheet.bis = CInt(.Cells(i, SNCol).value)
                                            Else
                                                outputLine = configLine.Titel & " : Angabe für Sheet ist kein Range"
                                                If awinSettings.englishLanguage Then
                                                    outputLine = configLine.Titel & " : this is no range"
                                                End If
                                                outputCollection.Add(outputLine)
                                            End If
                                            configLine.sheetDescript = CStr(.Cells(i, TabUCol).value)

                                            If configLine.cellrange Then
                                                Dim colrange As String = CStr(.Cells(i, SNCol).value)
                                                Dim hstr1() As String = Split(colrange, ":")
                                                If hstr1.Length = 2 Then
                                                    configLine.column.von = CInt(hstr1(0))
                                                    configLine.column.bis = CInt(hstr1(1))
                                                ElseIf hstr1.Length = 1 Then
                                                    configLine.row.von = CInt(.Cells(i, SNCol).value)
                                                    configLine.row.bis = CInt(.Cells(i, SNCol).value)
                                                Else
                                                    outputLine = configLine.Titel & " : Angabe ist kein Range"
                                                    If awinSettings.englishLanguage Then
                                                        outputLine = configLine.Titel & " : this is no range"
                                                    End If
                                                    outputCollection.Add(outputLine)
                                                End If
                                            Else
                                                configLine.column.von = CInt(.Cells(i, SNCol).value)
                                                configLine.column.bis = CInt(.Cells(i, SNCol).value)
                                            End If
                                            configLine.columnDescript = CStr(.Cells(i, SUCol).value)

                                            If configLine.cellrange Then
                                                Dim colrange As String = CStr(.Cells(i, ZNCol).value)
                                                Dim hstr2() As String = Split(colrange, ":")
                                                If hstr2.Length = 2 Then
                                                    configLine.row.von = CInt(hstr2(0))
                                                    configLine.row.bis = CInt(hstr2(1))
                                                ElseIf hstr2.Length = 1 Then
                                                    configLine.row.von = CInt(.Cells(i, ZNCol).value)
                                                    configLine.row.bis = CInt(.Cells(i, ZNCol).value)
                                                Else
                                                    outputLine = configLine.Titel & " : Angabe ist kein Range"
                                                    If awinSettings.englishLanguage Then
                                                        outputLine = configLine.Titel & " : this is no range"
                                                    End If
                                                    outputCollection.Add(outputLine)
                                                End If
                                            Else
                                                configLine.row.von = CInt(.Cells(i, ZNCol).value)
                                                configLine.row.bis = CInt(.Cells(i, ZNCol).value)
                                            End If
                                            configLine.rowDescript = CStr(.Cells(i, ZUCol).value)
                                            configLine.objType = CStr(.Cells(i, ObjCol).value)
                                            configLine.content = CStr(.Cells(i, InhaltCol).value)
                                    End Select

                                    If ActualDataConfigs.ContainsKey(configLine.Titel) Then
                                        ActualDataConfigs.Remove(configLine.Titel)
                                    End If

                                    ActualDataConfigs.Add(configLine.Titel, configLine)

                                Next

                            End With

                        End If

                    End If

                Catch ex As Exception
                    ' tk 5.2 es trat ein Fehler auf ... also Clear, weil das die ok / nicht ok Rückgabe Bedingung ist 
                    ActualDataConfigs.Clear()

                    If awinSettings.englishLanguage Then
                        outputLine = "The configurationfile " & configFile & " has no Sheet with name VISBO Config" & vbCrLf & " ... no import!"
                    Else
                        outputLine = "Die Konfigurationsdatei " & configFile & " enthält kein Registerblatt VISBO Config" &
                                    vbCrLf & " es fand kein Import statt "
                    End If
                    outputCollection.Add(outputLine)
                End Try

                ' configActualDataImport - Konfigurationsfile schließen
                configWB.Close(SaveChanges:=False)

            Catch ex As Exception
                ' tk 5.2 es trat ein Fehler auf ... also Clear, weil das die ok / nicht ok Rückgabe Bedingung ist 
                ActualDataConfigs.Clear()
                Call MsgBox("Das Öffnen der " & configFile & " war nicht erfolgreich")
            End Try

        End If

        checkActualDataImportConfig = (ActualDataConfigs.Count > 0)

    End Function

    ''' <summary>
    ''' überprüft, ob die Voraussetzungen für das Einlesen der Organisation
    ''' </summary>
    ''' <param name="configFile"></param>
    ''' <param name="orgaFile"></param>
    ''' <param name="orgaImportConfigs"></param>
    ''' <param name="lastrow"></param>
    ''' <returns></returns>
    Public Function checkOrgaImportConfig(ByVal configFile As String,
                                      ByRef orgaFile As String,
                                      ByRef orgaImportConfigs As SortedList(Of String, clsConfigOrgaImport),
                                      ByRef lastrow As Integer,
                                      ByRef outputCollection As Collection) As Boolean

        Dim configLine As New clsConfigOrgaImport
        Dim currentDirectoryName As String = requirementsOrdner
        Dim configWB As Microsoft.Office.Interop.Excel.Workbook = Nothing
        Dim currentWS As Microsoft.Office.Interop.Excel.Worksheet = Nothing
        Dim searcharea As Microsoft.Office.Interop.Excel.Range = Nothing
        Dim outputLine As String

        ''
        '' Config-file wird geöffnet
        ' Filename ggf. mit Directory erweitern
        configFile = My.Computer.FileSystem.CombinePath(currentDirectoryName, configFile)

        ' öffnen des Files 
        If My.Computer.FileSystem.FileExists(configFile) Then

            Try

                configWB = appInstance.Workbooks.Open(configFile)

                Try
                    If appInstance.Worksheets.Count > 0 Then

                        'currentWS = CType(appInstance.Worksheets(1), Global.Microsoft.Office.Interop.Excel.Worksheet)
                        currentWS = CType(configWB.Worksheets("VISBO Config"), Global.Microsoft.Office.Interop.Excel.Worksheet)

                        Dim titleCol As Integer,
                            IdentCol As Integer,
                            InputFileCol As Integer,
                            TypCol As Integer,
                            DatenCol As Integer,
                            TabUCol As Integer, TabNCol As Integer,
                            SUCol As Integer, SNCol As Integer,
                            ZUCol As Integer, ZNCol As Integer,
                            ObjCol As Integer,
                            InhaltCol As Integer

                        searcharea = currentWS.Rows(5)          ' Zeile 5 enthält die verschieden Configurationselemente

                        titleCol = searcharea.Find("Titel").Column
                        IdentCol = searcharea.Find("Identifier").Column
                        InputFileCol = searcharea.Find("InputFile").Column
                        TypCol = searcharea.Find("Typ").Column
                        DatenCol = searcharea.Find("Datenbereich").Column
                        TabUCol = searcharea.Find("Tabellen-Name").Column
                        TabNCol = searcharea.Find("Tabellen-Nummer").Column
                        SUCol = searcharea.Find("Spaltenüberschrift").Column
                        SNCol = searcharea.Find("Spalten-Nummer").Column
                        ZUCol = searcharea.Find("Zeilenbeschriftung").Column
                        ZNCol = searcharea.Find("Zeilen-Nummer").Column
                        ObjCol = searcharea.Find("Objekt-Typ").Column
                        InhaltCol = searcharea.Find("Inhalt").Column

                        Dim ok As Boolean = (titleCol + IdentCol + TypCol + DatenCol + SUCol + SNCol + ZUCol + ZNCol + ObjCol + InhaltCol > 13)

                        If ok Then
                            With currentWS
                                lastrow = .Cells(.Rows.Count, titleCol).end(Microsoft.Office.Interop.Excel.XlDirection.xlUp).row

                                For i = 6 To lastrow

                                    configLine = New clsConfigOrgaImport

                                    Dim Titel As String = CStr(.Cells(i, titleCol).value)

                                    Select Case Titel
                                        Case "DateiName"
                                            configLine.Titel = CStr(.Cells(i, titleCol).value)
                                            configLine.Inputfile = CStr(.Cells(i, InputFileCol).value)
                                            orgaFile = configLine.Inputfile



                                        Case Else
                                            configLine.Titel = CStr(.Cells(i, titleCol).value)
                                            configLine.Identifier = CStr(.Cells(i, IdentCol).value)
                                            configLine.Inputfile = CStr(.Cells(i, InputFileCol).value)
                                            configLine.Typ = CStr(.Cells(i, TypCol).value)
                                            configLine.cellrange = (CStr(.Cells(i, DatenCol).value) = "Range")
                                            'configLine.sheet = CInt(.Cells(i, TabNCol).value)
                                            Dim tabrange As String = CStr(.Cells(i, TabNCol).value)
                                            Dim hstr() As String = Split(tabrange, ":")
                                            If hstr.Length = 2 Then
                                                configLine.sheet.von = CInt(hstr(0))
                                                configLine.sheet.bis = CInt(hstr(1))
                                            ElseIf hstr.Length = 1 Then
                                                configLine.sheet.von = CInt(.Cells(i, SNCol).value)
                                                configLine.sheet.bis = CInt(.Cells(i, SNCol).value)
                                            Else
                                                outputLine = configLine.Titel & " : Angabe für Sheet ist kein Range"
                                                If awinSettings.englishLanguage Then
                                                    outputLine = configLine.Titel & " : this is no range"
                                                End If
                                                outputCollection.Add(outputLine)
                                            End If
                                            configLine.sheetDescript = CStr(.Cells(i, TabUCol).value)

                                            If configLine.cellrange Then
                                                Dim colrange As String = CStr(.Cells(i, SNCol).value)
                                                Dim hstr1() As String = Split(colrange, ":")
                                                If hstr1.Length = 2 Then
                                                    configLine.column.von = CInt(hstr1(0))
                                                    configLine.column.bis = CInt(hstr1(1))
                                                ElseIf hstr1.Length = 1 Then
                                                    configLine.row.von = CInt(.Cells(i, SNCol).value)
                                                    configLine.row.bis = CInt(.Cells(i, SNCol).value)
                                                Else
                                                    outputLine = configLine.Titel & " : Angabe ist kein Range"
                                                    If awinSettings.englishLanguage Then
                                                        outputLine = configLine.Titel & " : this is no range"
                                                    End If
                                                    outputCollection.Add(outputLine)
                                                End If
                                            Else
                                                configLine.column.von = CInt(.Cells(i, SNCol).value)
                                                configLine.column.bis = CInt(.Cells(i, SNCol).value)
                                            End If
                                            configLine.columnDescript = CStr(.Cells(i, SUCol).value)

                                            If configLine.cellrange Then
                                                Dim colrange As String = CStr(.Cells(i, ZNCol).value)
                                                Dim hstr2() As String = Split(colrange, ":")
                                                If hstr2.Length = 2 Then
                                                    configLine.row.von = CInt(hstr2(0))
                                                    configLine.row.bis = CInt(hstr2(1))
                                                ElseIf hstr2.Length = 1 Then
                                                    configLine.row.von = CInt(.Cells(i, ZNCol).value)
                                                    configLine.row.bis = CInt(.Cells(i, ZNCol).value)
                                                Else
                                                    outputLine = configLine.Titel & " : Angabe ist kein Range"
                                                    If awinSettings.englishLanguage Then
                                                        outputLine = configLine.Titel & " : this is no range"
                                                    End If
                                                    outputCollection.Add(outputLine)
                                                End If
                                            Else
                                                configLine.row.von = CInt(.Cells(i, ZNCol).value)
                                                configLine.row.bis = CInt(.Cells(i, ZNCol).value)
                                            End If
                                            configLine.rowDescript = CStr(.Cells(i, ZUCol).value)
                                            configLine.objType = CStr(.Cells(i, ObjCol).value)
                                            configLine.content = CStr(.Cells(i, InhaltCol).value)
                                    End Select

                                    If orgaImportConfigs.ContainsKey(configLine.Titel) Then
                                        orgaImportConfigs.Remove(configLine.Titel)
                                    End If

                                    orgaImportConfigs.Add(configLine.Titel, configLine)

                                Next

                            End With

                        End If

                    End If

                Catch ex As Exception
                    ' tk 5.2 es trat ein Fehler auf ... also Clear, weil das die ok / nicht ok Rückgabe Bedingung ist 
                    orgaImportConfigs.Clear()

                    If awinSettings.englishLanguage Then
                        outputLine = "The configurationfile " & configFile & " has no Sheet with name VISBO Config" & vbCrLf & " ... no import!"
                    Else
                        outputLine = "Die Konfigurationsdatei " & configFile & " enthält kein Registerblatt VISBO Config" &
                                    vbCrLf & " es fand kein Import statt "
                    End If
                    outputCollection.Add(outputLine)
                End Try

                ' configActualDataImport - Konfigurationsfile schließen
                configWB.Close(SaveChanges:=False)

            Catch ex As Exception
                ' tk 5.2 es trat ein Fehler auf ... also Clear, weil das die ok / nicht ok Rückgabe Bedingung ist 
                orgaImportConfigs.Clear()
                Call MsgBox("Das Öffnen der " & configFile & " war nicht erfolgreich")
            End Try

        End If

        checkOrgaImportConfig = (orgaImportConfigs.Count > 0)

    End Function

    ''' <summary>
    ''' 
    ''' </summary>
    ''' <param name="ActualDataConfig"></param>
    ''' <param name="tmpDatei"></param>
    ''' <param name="oPCollection"></param>
    ''' <returns></returns>
    Public Function readActualDataWithConfig(ByVal ActualDataConfig As SortedList(Of String, clsConfigActualDataImport),
                                             ByVal tmpDatei As String,
                                             ByVal IstDatenDate As Date,
                                             ByRef cacheProjekte As clsProjekteAlle,
                                             ByRef validProjectNames As SortedList(Of String, SortedList(Of String, Double())),
                                             ByRef projectRoleNames(,) As String,
                                             ByRef projectRoleValues(,,) As Double,
                                             ByRef updatedProjects As Integer,
                                             ByRef oPCollection As Collection) As Boolean

        Dim err As New clsErrorCodeMsg
        Dim outputline As String = ""
        Dim ok As Boolean = False
        Dim result As Boolean = True
        Dim actDataWB As Microsoft.Office.Interop.Excel.Workbook = Nothing
        Dim currentWS As Microsoft.Office.Interop.Excel.Worksheet = Nothing
        Dim regexpression As Regex
        Dim firstUrlTabelle As Integer
        Dim firstUrlspalte As Integer
        Dim firstUrlzeile As Integer
        Dim lastSpalte As Integer
        Dim lastZeile As Integer
        Dim hproj As New clsProjekt
        Dim pName As String = ""
        Dim anz_Proj As Integer = 0
        Dim searcharea As Microsoft.Office.Interop.Excel.Range = Nothing
        Dim t As Integer = 0  ' tabellenIndex
        Dim hrole As clsRollenDefinition = Nothing
        Dim personalName As String = ""
        Dim personalNumber As String = ""
        Dim curmonth As Integer
        Dim lastValidMonth As Integer = getColumnOfDate(IstDatenDate)
        Dim stundenTotal As Double = 0                     ' Stundenangabe in einer Zeile

        ' ======================
        ' vorarbeit der Definitionen geleistet
        ' ======================
        Try
            If My.Computer.FileSystem.FileExists(tmpDatei) Then
                Try
                    appInstance.DisplayAlerts = False
                    actDataWB = appInstance.Workbooks.Open(tmpDatei, UpdateLinks:=0)
                    actDataWB.Final = False
                    appInstance.DisplayAlerts = True

                    Dim vstart As clsConfigActualDataImport = ActualDataConfig("valueStart")
                    ' Auslesen erste Time-Sheet
                    firstUrlTabelle = vstart.sheet.von
                    firstUrlspalte = vstart.column.von
                    firstUrlzeile = vstart.row.von

                    ' Schleife über alle Tabellenblätter eines ausgewählten Excel-Files (hier = einer Rolle)
                    For t = 0 To vstart.sheet.bis - 1

                        If Not IsNothing(vstart.sheet.von + t) Then
                            currentWS = CType(actDataWB.Worksheets(vstart.sheet.von + t), Global.Microsoft.Office.Interop.Excel.Worksheet)
                            If Not IsNothing(vstart.sheetDescript) Then
                                ok = (vstart.sheetDescript.Contains(currentWS.Name))
                            Else
                                ok = True
                            End If
                        End If

                        If Not ok Then
                            If awinSettings.englishLanguage Then
                                outputline = "the sheet " & currentWS.Name & " doesn't match with the configuration"
                            Else
                                outputline = "das Tabellenblatt " & currentWS.Name & " passt nicht zur Konfiguration"
                            End If
                            oPCollection.Add(outputline)
                            Call logger(ptErrLevel.logError, outputline, "readActualDataWithConfig", anzFehler)
                            result = False
                            Exit For ' keine weiteren Tabellenblätter mehr lesen - Fehler aufgetreten
                        End If

                        If IsNothing(currentWS) Then
                            If awinSettings.englishLanguage Then
                                outputline = "the sheet " & vstart.sheetDescript & " doesn't exists in this workbook"
                            Else
                                outputline = "das Tabellenblatt " & vstart.sheetDescript & " ist nicht vorhanden"
                            End If
                            oPCollection.Add(outputline)
                            Call logger(ptErrLevel.logError, outputline, "readActualDataWithConfig", anzFehler)
                            result = False
                        Else
                            ' passendes Worksheet gefunden

                            Try
                                ' Find Month
                                Dim monat As String = currentWS.Cells(ActualDataConfig("months").row.von, ActualDataConfig("months").column.von).value
                                Dim vglMonat As String = currentWS.Name
                                Dim validm As Boolean = (vglMonat.Contains(monat) Or monat.Contains(vglMonat))
                                ' find Year
                                Dim jahr As String = currentWS.Cells(ActualDataConfig("years").row.von, ActualDataConfig("years").column.von).value
                                Dim vglJahr As String = currentWS.Name
                                Dim validj As Boolean = (vglJahr.Contains(jahr) Or jahr.Contains(vglJahr))
                                Dim xxx As Date = CDate("01." & monat & " " & jahr)
                                curmonth = getColumnOfDate(xxx)

                            Catch ex As Exception
                                outputline = "Error looking for month/year"
                                oPCollection.Add(outputline)
                                Call logger(ptErrLevel.logError, outputline, "readActualDataWithConfig", anzFehler)
                                result = False
                            End Try

                            If curmonth <= lastValidMonth Then

                                ' Find Wertespalte - auf jedem Tabellenblatt evt. anders
                                Dim hspalte As String = ActualDataConfig("Total").columnDescript
                                Dim stdSpalteTotal As Integer = 0
                                Try
                                    Dim überschriftenzeile As Integer = ActualDataConfig("Überschriften").row.von
                                    searcharea = currentWS.Rows(überschriftenzeile)          ' Zeile über... enthält die verschieden Spaltendescript
                                    stdSpalteTotal = searcharea.Find(hspalte).Column
                                Catch ex As Exception
                                    If awinSettings.englishLanguage Then
                                        outputline = "Error: in the sheet " & vstart.sheetDescript & " the value-column " & hspalte & " not found"
                                    Else
                                        outputline = "Error: im Tabellenblatt " & vstart.sheetDescript & " konnte die WerteSpalte " & hspalte & " nicht gefunden werden"
                                    End If
                                    oPCollection.Add(outputline)
                                    Call logger(ptErrLevel.logError, outputline, "readActualDataWithConfig", anzFehler)
                                    result = False
                                End Try

                                ' find PersoNr
                                Dim vPersoNr As clsConfigActualDataImport = ActualDataConfig("PersonalNumber")
                                Try
                                    personalNumber = currentWS.Cells(vPersoNr.row.von, vPersoNr.column.von).value
                                    ' find PersonalName
                                    Dim vPersoName As clsConfigActualDataImport = ActualDataConfig("PersonalName")
                                    personalName = currentWS.Cells(vPersoName.row.von, vPersoName.column.von).value
                                    hrole = RoleDefinitions.getRoledefByEmployeeNr(personalNumber)
                                    If IsNothing(hrole) Then
                                        ' Try Name 
                                        hrole = RoleDefinitions.getRoledef(personalName)
                                        If IsNothing(hrole) Then
                                            If awinSettings.englishLanguage Then
                                                outputline = "Person does not exist in organisation: '" & currentWS.Name & "' of File '" & tmpDatei & "' " & vbLf &
                                                    personalNumber & " : " & personalName
                                            Else
                                                outputline = "Person existiert nicht in Organisation: '" & currentWS.Name & "' in der Datei '" & tmpDatei & "' " & vbLf &
                                                    personalNumber & " : " & personalName
                                            End If
                                            oPCollection.Add(outputline)
                                            Call logger(ptErrLevel.logError, outputline, "readActualDataWithConfig", anzFehler)
                                            result = False
                                        Else
                                            If awinSettings.englishLanguage Then
                                                outputline = "Warning: Personell number does not match to Name '" & currentWS.Name & "' of File '" & tmpDatei & "' " & vbLf &
                                                personalNumber & " : " & personalName & " (Nr in VISBO: " & hrole.employeeNr & " )"
                                            Else
                                                outputline = "Warning: Personal Nummer passt nicht zu Name '" & currentWS.Name & "' in der Datei '" & tmpDatei & "' " & vbLf &
                                                personalNumber & " : " & personalName & " (Nr in VISBO: " & hrole.employeeNr & " )"
                                            End If

                                            Call logger(ptErrLevel.logError, outputline, "readActualDataWithConfig", anzFehler)
                                        End If

                                        'Call MsgBox(" hier ist der Fehler: " & personalNumber & ":" & personalName)
                                    Else
                                        If hrole.name <> personalName Then
                                            ' Warning: name and personal Number do not match ...
                                            If awinSettings.englishLanguage Then
                                                outputline = "Warning: Personell number does not match to Name '" & currentWS.Name & "' of File '" & tmpDatei & "' " & vbLf &
                                                personalNumber & " : " & personalName & " (Name in VISBO: " & hrole.name & " )"
                                            Else
                                                outputline = "Warning: Personal Nummer passt nicht zu Name '" & currentWS.Name & "' in der Datei '" & tmpDatei & "' " & vbLf &
                                                personalNumber & " : " & personalName & " (Name in VISBO: " & hrole.name & " )"
                                            End If

                                            Call logger(ptErrLevel.logError, outputline, "readActualDataWithConfig", anzFehler)
                                        End If
                                    End If
                                    'Dim identical As Boolean = (personalName = hrole.name)

                                Catch ex As Exception
                                    If awinSettings.englishLanguage Then
                                        outputline = "Error: in the sheet " & vstart.sheetDescript & "- there is something wrong with 'personal-No' or 'personal name'"
                                    Else
                                        outputline = "Fehler: im Tabellenblatt " & vstart.sheetDescript & "- es gibt ein Fehler beim lesen der Personalnummer oder des Namens"
                                    End If
                                    oPCollection.Add(outputline)
                                    Call logger(ptErrLevel.logError, outputline, "readActualDataWithConfig", anzFehler)
                                    result = False
                                End Try

                                lastSpalte = CType(currentWS.Cells(firstUrlzeile, 2000), Global.Microsoft.Office.Interop.Excel.Range).End(Excel.XlDirection.xlToLeft).Column
                                lastZeile = CType(currentWS.Cells(2000, firstUrlspalte), Global.Microsoft.Office.Interop.Excel.Range).End(Excel.XlDirection.xlUp).Row

                                If Not IsNothing(ActualDataConfig("valueEnd").rowDescript) Then
                                    Dim hzeile As String = ActualDataConfig("valueEnd").rowDescript
                                    Dim valueEndspalte As Integer = ActualDataConfig("valueEnd").column.von
                                    searcharea = currentWS.Columns(valueEndspalte)          ' in einer Spalte nach bestimmten Inhalt suchen
                                    lastZeile = searcharea.Find(hzeile).Row                 ' ZeilenNummer diesen Inhaltes merken
                                End If

                                If result = True Then
                                    ' alle Zeilen eines Tabellenblattes lesen
                                    For z = firstUrlzeile To lastZeile

                                        stundenTotal = 0                ' zurücksetzen

                                        ' find ProjectNumber and the relevant Project
                                        Dim projektKDNr As String = ""
                                        Dim projKDNrConfig As clsConfigActualDataImport = ActualDataConfig("ProjectNumber")
                                        projektKDNr = CStr(currentWS.Cells(z, projKDNrConfig.column.von).value)

                                        If Not IsNothing(projektKDNr) Then

                                            If projKDNrConfig.objType = "RegEx" Then
                                                If Not IsNothing(projKDNrConfig.content) Then
                                                    regexpression = New Regex(projKDNrConfig.content)
                                                    Dim match As Match = regexpression.Match(projektKDNr)
                                                    If match.Success Then
                                                        projektKDNr = match.Value
                                                    Else
                                                        projektKDNr = Nothing
                                                        If awinSettings.englishLanguage Then
                                                            outputline = "Attention: " & hrole.name & " Sheet: " & currentWS.Name & " Line: " & z.ToString & " no project No. given!"
                                                        Else
                                                            outputline = "Achtung: " & hrole.name & " Tabelle: " & currentWS.Name & " Zeile: " & z.ToString & " keine ProjektNr. angegeben!"
                                                        End If
                                                        oPCollection.Add(outputline)
                                                        Call logger(ptErrLevel.logWarning, outputline, "readActualDataWithConfig", anzFehler)
                                                    End If
                                                End If
                                            End If
                                        End If

                                        If Not IsNothing(projektKDNr) Then

                                            Dim projektName As String = ""
                                            projektName = CStr(currentWS.Cells(z, ActualDataConfig("ProjectName").column.von).value)

                                            stundenTotal = CDbl(currentWS.Cells(z, stdSpalteTotal).value)

                                            ' Check mit der Summenbildung in der Zeile
                                            ' die Werte gehen erst in der Spalte 6 los, also column.von + 4
                                            'Dim stdRange As Excel.Range = CType(currentWS.Range(currentWS.Cells(z, vstart.column.von + 2), currentWS.Cells(z, stdSpalteTotal - 2)), Microsoft.Office.Interop.Excel.Range)
                                            Dim stdRange As Excel.Range = CType(currentWS.Range(currentWS.Cells(z, vstart.column.von + 4), currentWS.Cells(z, stdSpalteTotal - 2)), Microsoft.Office.Interop.Excel.Range)
                                            Dim stundenSumme As Double = 0
                                            Try
                                                ' tk hat bei Matthias Urch mal Fehler produziert - deswegen Warnung ausgeben , aber weiter machen 
                                                ' da konnte die Worksheet Funktion .sum nicht ausgeführt werden 
                                                stundenSumme = appInstance.WorksheetFunction.Sum(stdRange)

                                                If stundenTotal <> stundenSumme Then
                                                    If awinSettings.englishLanguage Then
                                                        outputline = "Attention: " & hrole.name & ": in '" & currentWS.Name & "': sum of the single values (" & stundenSumme.ToString & ") isn 't the same as the value in column '" & hspalte & "' (" & stundenTotal.ToString & ")"
                                                    Else
                                                        outputline = "Achtung: " & hrole.name & ": in '" & currentWS.Name & "': Die Summe der einzelnen Werte (" & stundenSumme.ToString & ") ist nicht gleich dem Eintrag in Spalte '" & hspalte & "' (" & stundenTotal.ToString & ")"
                                                    End If
                                                    oPCollection.Add(outputline)
                                                    Call logger(ptErrLevel.logWarning, outputline, "readActualDataWithConfig", anzFehler)
                                                End If

                                            Catch ex As Exception

                                                stundenSumme = stundenTotal

                                                If awinSettings.englishLanguage Then
                                                    outputline = "Attention: " & hrole.name & ": " & ex.Message & currentWS.Name
                                                Else
                                                    outputline = "Achtung: " & hrole.name & ": " & ex.Message & currentWS.Name
                                                End If
                                                oPCollection.Add(outputline)
                                                Call logger(ptErrLevel.logWarning, outputline, "readActualDataWithConfig", anzFehler)
                                            End Try


                                            Dim pvkey As String
                                            If Not IsNothing(projektName) Then
                                                pvkey = calcProjektKey(projektName, "")
                                            Else
                                                pvkey = ""
                                            End If

                                            If cacheProjekte.containsPNr(projektKDNr) Then
                                                hproj = cacheProjekte.getProjectByKDNr(projektKDNr)
                                                pName = hproj.name
                                            Else
                                                hproj = Nothing         ' Vorbesetzung

                                                Dim pNames As Collection = CType(databaseAcc, DBAccLayer.Request).retrieveProjectNamesByPNRFromDB(projektKDNr, err)
                                                If pNames.Count = 1 Then
                                                    pName = pNames(1)

                                                    Dim pname_ok As Boolean = pName = projektName
                                                    ' Meldung noch ins Logbuch, wenn die Namen nicht übereinstimmen
                                                    If Not pname_ok Then
                                                        If awinSettings.englishLanguage Then
                                                            outputline = "different projectnames of project No. '" & projektKDNr & "': in the sheet it's called '" & projektName & "' in the DB it's called '" & pName & "'"
                                                        Else
                                                            outputline = "Unterschiedlicher Projektname für Projekt Nr. '" & projektKDNr & "': in der ExcelTabelle heißt es '" & projektName & "' in der DB  '" & pName & "'"
                                                        End If
                                                        Call logger(ptErrLevel.logWarning, outputline, "readActualDataWithConfig", anzFehler)
                                                    End If

                                                    hproj = New clsProjekt
                                                    hproj = CType(databaseAcc, DBAccLayer.Request).retrieveOneProjectfromDB(pName, "", "", Date.Now, err)

                                                ElseIf pNames.Count > 1 Then
                                                    ' Fehlermeldung, falls mehrer Projekte zu einer ProjektKdNr. existieren
                                                    If awinSettings.englishLanguage Then
                                                        outputline = "There exists more than one project to project No. '" & projektKDNr & "'"
                                                    Else
                                                        outputline = "Zu Projekt-Nr. '" & projektKDNr & "'" & " existieren mehrer Projekte"
                                                    End If

                                                    oPCollection.Add(outputline)
                                                    Call logger(ptErrLevel.logError, outputline, "readActualDataWithConfig", anzFehler)

                                                Else
                                                    ' Fehlermeldung, falls kein Projekt zu einer ProjektKdNr. existieren
                                                    If awinSettings.englishLanguage Then
                                                        outputline = "No project to project No. '" & projektKDNr & "' User: '" & hrole.name & "' month: '" & currentWS.Name & "'"
                                                    Else
                                                        outputline = "Es existiert kein Projekt zu Projekt-Nr. '" & projektKDNr & "' User: '" & hrole.name & "' Monat: '" & currentWS.Name & "'"
                                                    End If
                                                    oPCollection.Add(outputline)
                                                    Call logger(ptErrLevel.logError, outputline, "readActualDataWithConfig", anzFehler)

                                                End If
                                            End If

                                            If IsNothing(hproj) Then
                                                'Fehler, Projekt mit einer ProjektNr. existiert in DB nicht, Keine Istdaten hierzu einlesbar
                                                'If awinSettings.englishLanguage Then
                                                '    outputline = "project Nr. " & projektKDNr & " doesn't exist in the DB. No actual data can be stored"
                                                'Else
                                                '    outputline = "Projekt mit der  Projekt-Nummer " & projektKDNr & "existiert in der DB nicht. Istdaten sind nicht zuordenbar"
                                                'End If
                                                'oPCollection.Add(outputline)
                                                'Call logfileSchreiben(outputline, "readActualDataWithConfig", anzFehler)
                                                'result = False

                                            Else
                                                cacheProjekte.Add(hproj, updateCurrentConstellation:=False)                    ' Projekt in cacheProjekte merken

                                                Dim projBeginn = getColumnOfDate(hproj.startDate)
                                                Dim projEnde As Integer = getColumnOfDate(hproj.endeDate)

                                                ' Aufbauen des Eintrags
                                                Dim roleValues As New SortedList(Of String, Double())
                                                Dim tmpValues() As Double

                                                ReDim tmpValues(lastValidMonth - projBeginn)
                                                Dim teamID As Integer = -1

                                                If Not IsNothing(hrole) Then

                                                    Dim roleNameID As String = RoleDefinitions.bestimmeRoleNameID(hrole.name, "")

                                                    If Not validProjectNames.ContainsKey(pName) Then

                                                        roleValues = New SortedList(Of String, Double())

                                                        ' es handelt sich um Stunden, also in PT umrechnen 
                                                        tmpValues(curmonth - projBeginn) = stundenTotal / 8

                                                        roleValues.Add(roleNameID, tmpValues)
                                                        validProjectNames.Add(pName, roleValues)

                                                    Else
                                                        roleValues = validProjectNames.Item(pName)
                                                        If roleValues.ContainsKey(roleNameID) Then
                                                            ' rolle ist bereits enthalten 
                                                            ' also summieren 
                                                            tmpValues = roleValues.Item(roleNameID)

                                                            tmpValues(curmonth - projBeginn) = tmpValues(curmonth - projBeginn) + stundenTotal / 8


                                                        Else
                                                            ' Rolle ist noch nicht enthalten 

                                                            ' es handelt sich Stunden, also in PT umrechnen 
                                                            tmpValues(curmonth - projBeginn) = stundenTotal / 8

                                                            roleValues.Add(roleNameID, tmpValues)
                                                        End If

                                                    End If

                                                Else
                                                    'Fehler, darf nur ein Name zu einer ProjektNr. existieren => TimeSheets nicht ins archiv
                                                    If awinSettings.englishLanguage Then
                                                        outputline = "Role '" & hrole.name & "' does not exist in your organization"
                                                    Else
                                                        outputline = hrole.name & " ist nicht in Ihrer Organisation enthalten!"
                                                    End If

                                                    oPCollection.Add(outputline)
                                                    result = False
                                                End If
                                            End If
                                        Else
                                            'Fehler, es ist keine ProjektKDNr angegeben, Keine Istdaten hierzu einlesbar
                                            If stundenTotal <> 0 Then
                                                If awinSettings.englishLanguage Then
                                                    outputline = "Error: Actual Data cannot be imported: '" & hrole.name & "'/'" & currentWS.Name & "' There exists no project No. in line " & z.ToString
                                                Else
                                                    outputline = "Fehler: Istdaten sind nicht zuordenbar: '" & hrole.name & "'/'" & currentWS.Name & "' Es ist keine Projekt-Nummer angegeben in Zeile " & z.ToString
                                                End If
                                                oPCollection.Add(outputline)
                                                Call logger(ptErrLevel.logError, outputline, "readActualDataWithConfig", anzFehler)
                                                result = False
                                            End If
                                        End If      ' if ProjektKDNr = ""

                                    Next z          'nächste Zeile lesen
                                Else

                                End If

                            Else
                                ' Infomeldung im Logbuch
                                If awinSettings.englishLanguage Then
                                    outputline = "Finished  reading actual-data of " & personalName
                                Else
                                    outputline = "Ende der Istdaten für '" & personalName & "' erreicht"
                                End If

                                Call logger(ptErrLevel.logInfo, outputline, "readActualDataWithConfig", anzFehler)
                                Exit For
                            End If

                        End If

                    Next t    ' nächste Tabelle des Excel-Inputfiles

                Catch ex As Exception
                    actDataWB = Nothing
                    Call logger(ptErrLevel.logError, "1. " & ex.Message, "readActualDataWithConfig", anzFehler)
                    Call MsgBox("1. " & ex.Message)
                End Try

                If Not IsNothing(actDataWB) Then
                    actDataWB.Close(SaveChanges:=False)
                End If


            End If
        Catch ex As Exception
            Call logger(ptErrLevel.logError, "2. " & ex.Message, "readActualDataWithConfig", anzFehler)
            Call MsgBox("2. " & ex.Message)
        End Try

        readActualDataWithConfig = result
    End Function

    Public Function readActualData(ByVal dateiname As String) As Boolean

        'dateiname = My.Computer.FileSystem.CombinePath(dirname, selectedWB)

        Dim oCollection As New Collection

        Try
            ' hier wird jetzt der Import gemacht 
            Call logger(ptErrLevel.logInfo, "Beginn Import Istdaten", dateiname, -1)

            ' Öffnen des Organisations-Files
            appInstance.Workbooks.Open(dateiname)
            Dim scenarioNameP As String = appInstance.ActiveWorkbook.Name



            ' das Formular aufschalten mit 
            '
            'Dim editActualDataMonth As New frmProvideActualDataMonth

            'If editActualDataMonth.ShowDialog = DialogResult.OK Then

            '    Dim monat As Integer = CInt(editActualDataMonth.valueMonth.Text)

            '    Dim readPastAndFutureData As Boolean = editActualDataMonth.readPastAndFutureData.Checked
            '    Dim createUnknownProjects As Boolean = editActualDataMonth.createUnknownProjects.Checked


            '    Call ImportIstdatenStdFormat(monat, readPastAndFutureData, createUnknownProjects, oCollection)

            'End If

            Dim readAll As Boolean = False
            Call ImportIstdatenStdFormat(readAll, oCollection)

            Dim wbName As String = My.Computer.FileSystem.GetName(dateiname)

            ' Schliessen des CustomUser Role-Files
            appInstance.Workbooks(wbName).Close(SaveChanges:=True)

            If oCollection.Count = 0 Then
                'sessionConstellationP enthält alle Projekte aus dem Import 
                Dim sessionConstellationP As clsConstellation = verarbeiteImportProjekte(scenarioNameP, noComparison:=False, considerSummaryProjects:=False)

                Dim scenarioPVName As String = calcPortfolioKey(scenarioNameP, "")
                If sessionConstellationP.count > 0 Then

                    If projectConstellations.Contains(scenarioPVName) Then
                        projectConstellations.Remove(scenarioPVName)
                    End If

                    projectConstellations.Add(sessionConstellationP)
                    ' jetzt auf Projekt-Tafel anzeigen 
                    Call loadSessionConstellation(sessionConstellationP.constellationName, False, True)

                Else
                    Call MsgBox("keine Projekte importiert ...")
                End If

                If ImportProjekte.Count > 0 Then
                    ImportProjekte.Clear(False)
                End If
            Else
                Call showOutPut(oCollection, "Errors occurred .. no import", "")
            End If


        Catch ex As Exception

        End Try


        readActualData = (oCollection.Count = 0)
    End Function

    ''' <summary>
    ''' liest das im Diretory ../ressource manager evt. liegende File 'Urlaubsplaner*.xlsx' File  aus
    ''' und hinterlegt an entsprechender Stelle im hrole.kapazitaet die verfügbaren Tage der entsprechenden Rolle
    ''' </summary>
    ''' <remarks></remarks>
    Friend Function readAvailabilityOfRole(ByVal kapaFileName As String, ByRef oPCollection As Collection) As Boolean

        Dim err As New clsErrorCodeMsg
        Dim old_oPCollectionCount As Integer = oPCollection.Count

        Dim ok As Boolean = True
        Dim formerEE As Boolean = appInstance.EnableEvents
        Dim formerSU As Boolean = appInstance.ScreenUpdating
        Dim msgtxt As String = ""
        Dim anzFehler As Integer = 0
        Dim fehler As Boolean = False

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


                            lastSpalte = CType(currentWS.Cells(4, 2000), Global.Microsoft.Office.Interop.Excel.Range).End(Excel.XlDirection.xlToLeft).Column
                            lastZeile = CType(currentWS.Cells(2000, 1), Global.Microsoft.Office.Interop.Excel.Range).End(Excel.XlDirection.xlUp).Row

                            ' letzte Zeile bestimmen, wenn dies verbunden Zellen sind
                            ' -------------------------------------
                            Dim rng As Range
                            Dim rngEnd As Range

                            rng = CType(currentWS.Cells(lastZeile, 1), Global.Microsoft.Office.Interop.Excel.Range)

                            If rng.MergeCells Then

                                rng = rng.MergeArea
                                rngEnd = rng.Cells(rng.Rows.Count, rng.Columns.Count)

                                ' dann ist die lastZeile neu zu besetzen
                                lastZeile = rngEnd.Row
                            End If

                            ' nun hat die Variable lastZeile sicher den richtigen Wert
                            ' --------------------------------------


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
                                    msgtxt = "Error reading planning holidays: Please check the calendar in this file ..."
                                Else
                                    msgtxt = "Fehler beim Lesen der Urlaubsplanung: Bitte prüfen Sie die Korrektheit des Kalenders ..."
                                End If
                                If Not oPCollection.Contains(msgtxt) Then
                                    oPCollection.Add(msgtxt, msgtxt)
                                End If
                                'Call MsgBox(msgtxt)

                                Call logger(ptErrLevel.logError, msgtxt, kapaFileName, anzFehler)

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

                                Call logger(ptErrLevel.logWarning, msgtxt, kapaFileName, anzFehler)
                                'Call showOutPut(oPCollection, "Lesen Urlaubsplanung wurde mit Fehler abgeschlossen", "Meldungen zu Lesen Urlaubsplanung")
                                ' tk 12.2.19 ess oll alles gelesen werden - es wird nicht weitergemacht, wenn es Einträge in der outputCollection gibt 
                                'Throw New ArgumentException(msgtxt)
                            Else

                                For iZ = 5 To lastZeile


                                    rolename = CType(currentWS.Cells(iZ, 2), Global.Microsoft.Office.Interop.Excel.Range).Text
                                    If rolename <> "" Then
                                        hrole = RoleDefinitions.getRoledef(rolename)
                                        If Not IsNothing(hrole) Then

                                            Dim defaultHrsPerdayForThisPerson As Double = hrole.defaultDayCapa

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

                                                                If IsNumeric(CType(currentWS.Cells(iZ, sp), Global.Microsoft.Office.Interop.Excel.Range).Value) Then

                                                                    Dim angabeInStd As Double = CType(CType(currentWS.Cells(iZ, sp), Global.Microsoft.Office.Interop.Excel.Range).Value, Double)

                                                                    If angabeInStd >= 0 And angabeInStd <= 24 Then
                                                                        anzArbStd = anzArbStd + CDbl(CType(currentWS.Cells(iZ, sp), Global.Microsoft.Office.Interop.Excel.Range).Value)
                                                                    Else
                                                                        If awinSettings.englishLanguage Then
                                                                            msgtxt = "Error reading the amount of working hours for " & hrole.name & " : " & angabeInStd.ToString & " (!!)"
                                                                        Else
                                                                            msgtxt = "Fehler beim Lesen der Anzahl zu leistenden Arbeitsstunden " & hrole.name & " : " & angabeInStd.ToString & " (!!)"
                                                                        End If
                                                                        If Not oPCollection.Contains(msgtxt) Then
                                                                            oPCollection.Add(msgtxt, msgtxt)
                                                                        End If
                                                                        'Call MsgBox(msgtxt)
                                                                        fehler = True
                                                                        Call logger(ptErrLevel.logError, msgtxt, kapaFileName, anzFehler)
                                                                    End If
                                                                Else
                                                                    ' Feld ist weiss, oder hat keine Farbe, keine Zahl: also ist es Arbeitstag mit Default-Std pro Tag 
                                                                    anzArbStd = anzArbStd + defaultHrsPerdayForThisPerson
                                                                End If



                                                            Else

                                                                ' hier wird die Telair Variante gemacht 
                                                                ' das einfachste wäre eigentlich  
                                                                'anzArbStd = anzArbStd + defaultHrsPerdayForThisPerson

                                                                Dim colorIndup As Integer = CType(currentWS.Cells(iZ, sp), Global.Microsoft.Office.Interop.Excel.Range).Borders(XlBordersIndex.xlDiagonalUp).ColorIndex

                                                                ' Wenn das Feld nicht durch einen Diagonalen Strich gekennzeichnet ist
                                                                If CType(currentWS.Cells(iZ, sp), Global.Microsoft.Office.Interop.Excel.Range).Borders(XlBordersIndex.xlDiagonalUp).ColorIndex = noColor Then
                                                                    'anzArbStd = anzArbStd + 8
                                                                    anzArbStd = anzArbStd + defaultHrsPerdayForThisPerson
                                                                Else
                                                                    ' freier Tag für Teilzeitbeschäftigte
                                                                    msgtxt = "Tag zählt nicht: Zeile " & iZ & ", Spalte " & sp
                                                                    Call logger(ptErrLevel.logInfo, msgtxt, kapaFileName, anzFehler)
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
                                                        Call logger(ptErrLevel.logError, msgtxt, kapaFileName, anzFehler)
                                                    End If

                                                Next

                                                anzArbTage = anzArbStd / 8

                                                'nur wenn die hrole schon eingetreten und nicht ausgetreten ist, wird die Capa eingetragen
                                                If colOfDate >= getColumnOfDate(hrole.entryDate) And colOfDate < getColumnOfDate(hrole.exitDate) Then
                                                    hrole.kapazitaet(colOfDate) = anzArbTage
                                                Else
                                                    hrole.kapazitaet(colOfDate) = 0
                                                End If

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
                                            Call logger(ptErrLevel.logError, msgtxt, kapaFileName, anzFehler)
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
                                        Call logger(ptErrLevel.logError, msgtxt, kapaFileName, anzFehler)
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
                                Call logger(ptErrLevel.logError, msgtxt, kapaFileName, anzFehler)
                            End If

                        End If

                    Next index


                Catch ex2 As Exception
                    'If fehler Then
                    '    'Call MsgBox(msgtxt)

                    '    RoleDefinitions = CType(databaseAcc, DBAccLayer.Request).retrieveRolesFromDB(DateTime.Now, err)

                    '    msgtxt = "Es wurden nun die Kapazitäten aus der Datenbank gelesen ..."
                    '    If awinSettings.englishLanguage Then
                    '        msgtxt = "Therefore read the capacity of every Role from the DB  ..."
                    '    End If
                    '    If Not oPCollection.Contains(msgtxt) Then
                    '        oPCollection.Add(msgtxt, msgtxt)
                    '    End If
                    '    Call logfileSchreiben(msgtxt, kapaFileName, anzFehler)
                    'End If
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

        ' das wird jetzt an der übergeordneten Stelle gemacht
        'Call showOutPut(oPCollection, "Meldungen zu Lesen Urlaubsplanung", "Folgende Probleme sind beim Lesen der Urlaubsplanung aufgetreten")

        ' ''If outPutCollection.Count > 0 Then
        ' ''    Call showOutPut(outPutCollection, _
        ' ''                    "Meldungen Einlesevorgang Urlaubsdatei", _
        ' ''                    "zum Zeitpunkt " & storedAtOrBefore.ToString & " aufgeführte Rolle nicht definiert")
        ' ''End If

        readAvailabilityOfRole = (oPCollection.Count = old_oPCollectionCount)

    End Function

    ''' <summary>
    ''' Berechnung der Anzahl Arbeitstage im AktMonat/AktJahr
    ''' </summary>
    ''' <param name="AktJahr"></param>
    ''' <param name="AktMonat"></param>
    ''' <returns></returns>
    Private Function WorkingDaysInMonth(ByVal AktJahr, ByVal AktMonat) As Integer
        Dim AnzahlTage As Integer = DateTime.DaysInMonth(AktJahr, AktMonat)
        Dim AnzahlArbeitsTage As Integer = 0
        For i As Integer = 1 To AnzahlTage
            Dim day As New Date(AktJahr, AktMonat, i)
            If Not (day.DayOfWeek = DayOfWeek.Sunday Or day.DayOfWeek = DayOfWeek.Saturday) Then
                AnzahlArbeitsTage += 1
            End If
        Next
        WorkingDaysInMonth = AnzahlArbeitsTage
    End Function
    Private Function freeDaysInMonth(ByVal AktJahr, ByVal AktMonat) As Integer
        Dim freeListe As New SortedList(Of Date, String)
        Dim NameFeiertag As String = officialHoliday(DateSerial(AktJahr, AktMonat, 1), freeListe)
        Dim AnzahlTage As Integer = DateTime.DaysInMonth(AktJahr, AktMonat)
        Dim AnzahlfreieTage As Integer = 0
        For i As Integer = 1 To AnzahlTage
            Dim day As Date = DateSerial(AktJahr, AktMonat, i)
            If freeListe.ContainsKey(day) Or day.DayOfWeek = DayOfWeek.Sunday Or day.DayOfWeek = DayOfWeek.Saturday Then
                AnzahlfreieTage += 1
            End If
        Next
        freeDaysInMonth = AnzahlfreieTage
    End Function
    ''' <summary>
    ''' erstellt einen Kalender, der Ausgangsbasis für Kapazitäten ist
    ''' </summary>
    ''' <returns></returns>
    Private Function createDefaultCalendar() As clsDefaultCalendar
        Dim defaultCal As New clsDefaultCalendar
        Dim monthCal As New clsBusinessDays
        Dim relMonth As Integer = getColumnOfDate(StartofCalendar)
        For y As Integer = Year(StartofCalendar) To Year(StartofCalendar) + 20 - 1
            For m As Integer = Month(StartofCalendar) To 12
                monthCal = New clsBusinessDays
                monthCal.year = y
                monthCal.month = m
                monthCal.noOfNonBusinessDays = freeDaysInMonth(y, m)
                monthCal.noOfBusinessDays = DateTime.DaysInMonth(y, m) - monthCal.noOfNonBusinessDays
                Dim check As Boolean = (monthCal.noOfBusinessDays = WorkingDaysInMonth(y, m))
                defaultCal.defCal.Add(relMonth, monthCal)
                relMonth += 1

            Next        ' for m 
        Next            ' for y
        createDefaultCalendar = defaultCal
    End Function



    ''' <summary>
    ''' liest das im Diretory ../ressource manager evt. liegende File 'zeuss*.xlsx' (oder wie in kapaConfig benamst) File  aus
    ''' und hinterlegt an entsprechender Stelle im hrole.kapazitaet die verfügbaren Tage der entsprechenden Rolle
    ''' </summary>
    ''' <remarks></remarks>
    Friend Function readAvailabilityOfRoleWithConfig(ByVal kapaConfig As SortedList(Of String, clsConfigKapaImport),
                                                ByVal kapaFileName As String,
                                                ByRef oPCollection As Collection) As Boolean

        Dim err As New clsErrorCodeMsg
        Dim old_oPCollectionCount As Integer = oPCollection.Count
        Dim kapaWB As Microsoft.Office.Interop.Excel.Workbook = Nothing

        Dim ok As Boolean = True
        Dim formerEE As Boolean = appInstance.EnableEvents
        Dim formerSU As Boolean = appInstance.ScreenUpdating
        Dim msgtxt As String = ""
        Dim anzFehler As Integer = 0
        Dim fehler As Boolean = False

        Dim ImportTyp As Integer = 2            ' Import like Telair - Zeuss - Dateien
        Try
            ImportTyp = kapaConfig("ImportTyp").content
        Catch ex As Exception
            '' Einlesen Kapa wie Telair - Zeuss
            'If awinSettings.englishLanguage Then
            '    Call MsgBox("ConfigFile with errors - Abort")
            'Else
            '    Call MsgBox("Fehlerhafte Konfigurationsdatei - Abbruch")
            'End If
        End Try

        Select Case ImportTyp
            Case 1            ' Import like Instart *Holidays*

            Case 2            ' Import like Telair - Zeuss - Dateien

            Case Else

        End Select


        If ImportTyp = 1 Then               ' Import like Instart *Holidays*

            ' zunächst den Default-Kalender ( von StartOfCalendar an 240 Monate) erstellen unter Berücksichtigung der Feiertage
            Dim defaultCal As clsDefaultCalendar = createDefaultCalendar()

            ' Read capacities and/or holidays for every role 
            Dim addOnHolidays As New SortedList(Of String, clsDefaultCalendar)

            Dim firstColumn As Integer = 0
            Dim firstRow As Integer = 0
            Dim lastRow As Integer = 0
            Dim lastColumn As Integer = 0

            'Dim noColor As Integer = -4142
            'Dim whiteColor As Integer = 2
            Dim currentWS As Excel.Worksheet
            Dim index As Integer = 1

            'Dim year As Integer = DatePart(DateInterval.Year, Date.Now)
            Dim monthName As String = ""

            Dim colDate As Integer = 0
            Dim anzDays As Integer = 0

            'Dim monthDays As New SortedList(Of Integer, Integer)

            Dim hrole As New clsRollenDefinition
            Dim rolename As String = ""
            Dim absenceDay As Date
            Dim absenceType As String = ""
            Dim input_ok As Boolean = True
            Dim regexpression As Regex

            Dim outPutCollection As New Collection

            If formerEE Then
                appInstance.EnableEvents = False
            End If

            If formerSU Then
                appInstance.ScreenUpdating = False
            End If

            enableOnUpdate = False

            Dim roleCol As Integer = kapaConfig("role").column
            Dim dateCol As Integer = kapaConfig("date").column
            Dim absenceCol As Integer = kapaConfig("absence type").column
            Dim roleBusy As New clsBusinessDays
            Dim roleCapa As New clsDefaultCalendar

            ' öffnen des Files 
            If My.Computer.FileSystem.FileExists(kapaFileName) Then

                Try
                    kapaWB = appInstance.Workbooks.Open(kapaFileName)

                    Try
                        For index = 1 To appInstance.Worksheets.Count

                            Dim tab As String = kapaConfig("valueStart").tabName

                            currentWS = CType(appInstance.Worksheets(index), Global.Microsoft.Office.Interop.Excel.Worksheet)

                            With currentWS

                                ' Auslesen erste Verfügbarkeitsspalte
                                firstColumn = kapaConfig("valueStart").column
                                firstRow = kapaConfig("valueStart").row
                                Dim lastLineConfig As String = kapaConfig("LastLine").content
                                If lastLineConfig = "" Then
                                    lastRow = CType(currentWS.Cells(10000, 1), Global.Microsoft.Office.Interop.Excel.Range).End(Excel.XlDirection.xlUp).Row
                                Else
                                    ' TODO: muss gemäss RegEx berechnet werden
                                End If

                                ' loop über die Zeilen
                                For ix As Integer = firstRow To lastRow

                                    input_ok = True                   ' Initialise
                                    rolename = CType(currentWS.Cells(ix, roleCol).value, String).Trim
                                    If IsNothing(rolename) Then
                                        input_ok = False
                                    End If

                                    absenceDay = CDate(currentWS.Cells(ix, dateCol).value)
                                    If IsNothing(absenceDay) Then
                                        input_ok = False
                                    End If

                                    absenceType = CStr(currentWS.Cells(ix, absenceCol).value)
                                    If IsNothing(absenceType) Then
                                        input_ok = False
                                    Else
                                        If kapaConfig("absence type").regex = "RegEx" Then
                                            'regexpression = New Regex("[0-9]{4}")
                                            regexpression = New Regex(kapaConfig("absence type").content)
                                            Dim match As Match = regexpression.Match(absenceType)
                                            If match.Success Then
                                                absenceType = match.Value
                                            Else
                                                absenceType = ""
                                                input_ok = False
                                            End If
                                        End If
                                    End If

                                    If input_ok Then        ' alle drei Angabe dieser Zeile sind soweit passend

                                        Dim columnOfDate As Integer = getColumnOfDate(absenceDay)
                                        If addOnHolidays.ContainsKey(rolename) Then
                                            roleCapa = addOnHolidays(rolename)
                                        Else
                                            roleCapa = New clsDefaultCalendar
                                        End If

                                        If roleCapa.defCal.ContainsKey(columnOfDate) Then
                                            roleBusy = roleCapa.defCal(columnOfDate)
                                        Else
                                            roleBusy = New clsBusinessDays
                                        End If

                                        roleBusy.month = Month(absenceDay)
                                        roleBusy.year = Year(absenceDay)

                                        Dim NameFeiertag As String = officialHoliday(absenceDay)
                                        If (NameFeiertag = "") _
                                                And Not (absenceDay.DayOfWeek = DayOfWeek.Sunday) _
                                                And Not (absenceDay.DayOfWeek = DayOfWeek.Saturday) Then
                                            ' absenceDay ist als zusätzlicher Nicht-Arbeitstag zu berücksichtigen
                                            roleBusy.noOfNonBusinessDays += 1
                                        Else
                                            ' Tag schon als nicht Arbeitstag berücksichtigt
                                        End If

                                        If roleCapa.defCal.ContainsKey(columnOfDate) Then
                                            roleCapa.defCal.Remove(columnOfDate)
                                        End If
                                        roleCapa.defCal.Add(columnOfDate, roleBusy)

                                        If addOnHolidays.ContainsKey(rolename) Then
                                            addOnHolidays.Remove(rolename)
                                        End If
                                        addOnHolidays.Add(rolename, roleCapa)

                                    Else
                                        If Not IsNothing(absenceType) And Not absenceType = "" Then
                                            If awinSettings.englishLanguage Then
                                                msgtxt = "Error in Line: " & ix & " not matching input " & vbLf & kapaFileName
                                            Else
                                                msgtxt = "Fehler in Zeile: " & ix & " Input passt nicht zusammen " & vbLf & kapaFileName
                                            End If
                                            'oPCollection.Add(msgtxt)
                                            Call logger(ptErrLevel.logError, msgtxt, kapaFileName, anzFehler)
                                        Else
                                            ' Zeile überlesen ohne Fehlermeldung
                                            Dim a As Integer = 0
                                        End If

                                    End If

                                Next   ' row

                            End With
                            'End If
                        Next
                    Catch ex As Exception

                    End Try

                Catch ex As Exception

                End Try

                ' Übertragen der Urlaubstage in die Kapazität der Organisations-mitglieder

                For Each kvp As KeyValuePair(Of String, clsDefaultCalendar) In addOnHolidays
                    rolename = kvp.Key
                    roleCapa = kvp.Value

                    ' bereits in orga vorhandene Kapa holen
                    hrole = RoleDefinitions.getRoledef(rolename)
                    If Not IsNothing(hrole) Then
                        For Each kvpCapa As KeyValuePair(Of Integer, clsBusinessDays) In roleCapa.defCal
                            ' default BusinessDays im Monat kvpCapa.key
                            Dim colofDate As Integer = kvpCapa.Key
                            ' Anzahl Arbeitstage, errechnet gemäß DefaultKalender
                            Dim defaultDays As Integer = defaultCal.defCal(kvpCapa.Key).noOfBusinessDays
                            ' nur die Tage, die kein Feiertag und kein WE sind
                            Dim Urlaubsdays As Integer = kvpCapa.Value.noOfNonBusinessDays
                            Dim anzArbTage As Integer = defaultDays - Urlaubsdays
                            ' capa = Kapazität, die für Projektarbeit bleibt
                            Dim capa As Double = anzArbTage * hrole.defaultDayCapa / 8
                            'nur wenn die hrole schon eingetreten und nicht ausgetreten ist, wird die Capa eingetragen
                            If colofDate >= getColumnOfDate(hrole.entryDate) And
                                colofDate < getColumnOfDate(hrole.exitDate) Then

                                hrole.kapazitaet(colofDate) = capa
                            Else
                                hrole.kapazitaet(colofDate) = 0
                            End If
                        Next
                    Else
                        If awinSettings.englishLanguage Then
                            msgtxt = "Warning: the role: " & rolename & " isn't defined in the Organisation " & vbLf & kapaFileName
                        Else
                            msgtxt = "Warning: die Person: " & rolename & " ist nicht in der Organisation enthalten " & vbLf & kapaFileName
                        End If
                        'oPCollection.Add(msgtxt)
                        Call logger(ptErrLevel.logWarning, msgtxt, kapaFileName, anzFehler)
                    End If

                Next

                Dim halt As Boolean = True

            End If


        ElseIf ImportTyp = 2 Then

            Dim spalte As Integer = 2
            Dim firstUrlspalte As Integer = 0
            Dim firstUrlzeile As Integer = 0
            Dim noColor As Integer = -4142
            Dim whiteColor As Integer = 2
            Dim currentWS As Excel.Worksheet
            Dim index As Integer
            Dim dateConsidered As Date

            'Dim year As Integer = DatePart(DateInterval.Year, Date.Now)
            Dim monthName As String = ""

            ' tk wird nicht verwendet ... 
            'Dim monthNumber As Integer = 0

            Dim Jahr As Integer = 0
            Dim anzMonthDays As Integer = 0
            Dim colDate As Integer = 0
            Dim anzDays As Integer = 0

            Dim lastZeile As Integer
            Dim lastSpalte As Integer
            Dim monthDays As New SortedList(Of Integer, Integer)

            Dim hrole As New clsRollenDefinition
            Dim rolename As String = ""

            Dim regexpression As Regex

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

                            currentWS = CType(appInstance.Worksheets(index), Global.Microsoft.Office.Interop.Excel.Worksheet)
                            With currentWS

                                'Dim regex As String = kapaConfig("month").regex
                                'Dim Inhalt As String = kapaConfig("month").content

                                ' Auslesen der Jahreszahl, falls vorhanden
                                Dim hjahr As String = CStr(.Cells(kapaConfig("year").row, kapaConfig("year").column).value)
                                If IsNothing(hjahr) Then
                                    Jahr = 0
                                Else
                                    If kapaConfig("year").regex = "RegEx" Then
                                        'regexpression = New Regex("[0-9]{4}")
                                        regexpression = New Regex(kapaConfig("year").content)
                                        Dim match As Match = regexpression.Match(hjahr)
                                        If match.Success Then
                                            Jahr = CInt(match.Value)
                                        Else
                                            Jahr = 0
                                        End If
                                    End If
                                End If


                                ' Auslesen des relevanten Monats
                                Dim hmonth As String = CStr(.Cells(kapaConfig("month").row, kapaConfig("month").column).value)
                                If IsNothing(hmonth) Then
                                    monthName = ""
                                Else
                                    If kapaConfig("month").regex = "RegEx" Then
                                        regexpression = New Regex(kapaConfig("month").content)
                                        Dim Match As Match = regexpression.Match(hmonth)
                                        If Match.Success Then
                                            monthName = Match.Value
                                        Else
                                            monthName = ""
                                        End If
                                    End If
                                End If


                                ' Auslesen erste Verfügbarkeitsspalte
                                firstUrlspalte = kapaConfig("valueStart").column
                                firstUrlzeile = kapaConfig("valueStart").row
                            End With

                            ' tk 3.2.20 
                            Dim isdate As Boolean = DateTime.TryParse(monthName & " " & Jahr.ToString, dateConsidered)

                            Dim beginningDay As Integer = -1
                            Dim endingDay As Integer = -1
                            Try
                                ' das kann schiefgehen, wenn keine Zahl im Feld steht ... 
                                beginningDay = CInt(currentWS.Cells(firstUrlzeile - 1, firstUrlspalte).value)
                            Catch ex As Exception
                                beginningDay = -1
                            End Try

                            If beginningDay <> 1 Then
                                If awinSettings.englishLanguage Then
                                    msgtxt = "Error in date definition row in Capa file: File, Row, Column: " & vbLf & kapaFileName & ", " & firstUrlzeile & ", " & firstUrlspalte & " does not start with 1"
                                Else
                                    msgtxt = "Fehler in Datums-Zeile in Kapazitäts-Datei: Datei, Zeile, Spalte: " & vbLf & kapaFileName & ", " & firstUrlzeile & ", " & firstUrlspalte & "startet nicht bei 1"

                                End If
                                oPCollection.Add(msgtxt)
                                Call logger(ptErrLevel.logError, msgtxt, kapaFileName, anzFehler)

                            ElseIf Not isdate Then

                                If awinSettings.englishLanguage Then
                                    msgtxt = "Error in Month of capacity definition: no valid month, year in Capa file: " & kapaFileName
                                Else
                                    msgtxt = "Fehler in Angabe des auszulesenden Monats in Kapazitäts-Datei: " & kapaFileName

                                End If
                                oPCollection.Add(msgtxt)
                                Call logger(ptErrLevel.logError, msgtxt, kapaFileName, anzFehler)
                            Else
                                If Jahr <> 0 And monthName <> "" Then

                                    colDate = getColumnOfDate(dateConsidered)

                                    monthDays.Clear()

                                    anzMonthDays = DateTime.DaysInMonth(Jahr, Month(dateConsidered))
                                    If Not monthDays.ContainsKey(colDate) Then
                                        monthDays.Add(colDate, anzMonthDays)
                                    End If

                                    ' tk prüfen, ob der letzte Tag auch der richtige ist ... 
                                    Try
                                        ' das kann schiefgehen, wenn keine Zahl im Feld steht ... 
                                        endingDay = CInt(currentWS.Cells(firstUrlzeile - 1, firstUrlspalte + anzMonthDays - 1).value)
                                    Catch ex As Exception
                                        endingDay = -1
                                    End Try

                                    If endingDay <> anzMonthDays Then
                                        ok = False

                                        If awinSettings.englishLanguage Then
                                            msgtxt = "Error in date definition row in Capa file: File, Row, Column: " & vbLf & kapaFileName & ", " & firstUrlzeile & ", " & firstUrlspalte + anzMonthDays - 1 & " does not show last day in month"
                                        Else
                                            msgtxt = "Fehler in Datums-Zeile in Kapazitäts-Datei: Datei, Zeile, Spalte: " & vbLf & kapaFileName & ", " & firstUrlzeile & ", " & firstUrlspalte + anzMonthDays - 1 & "zeigt nicht den letzten Tag des Monats "

                                        End If

                                        oPCollection.Add(msgtxt)
                                        Call logger(ptErrLevel.logError, msgtxt, kapaFileName, anzFehler)

                                    Else
                                        ' hier ist sichergestellt, dass die erste Spalte mit 1 beginnt, die letzte Spalte dem Tag entspricht, mit dem der Monat endet

                                        ok = True

                                        anzDays = 0

                                        lastSpalte = CType(currentWS.Cells(firstUrlzeile, 2000), Global.Microsoft.Office.Interop.Excel.Range).End(Excel.XlDirection.xlToLeft).Column
                                        lastZeile = CType(currentWS.Cells(2000, 1), Global.Microsoft.Office.Interop.Excel.Range).End(Excel.XlDirection.xlUp).Row

                                        ' Nachkorrektur gemäss Angabe in KonfigDate 'LastLine'
                                        Dim found As Boolean = False
                                        Dim i As Integer = lastZeile + 1
                                        While Not found
                                            i = i - 1
                                            If kapaConfig("LastLine").regex = "RegEx" Then
                                                regexpression = New Regex(kapaConfig("LastLine").content)
                                                Dim lastLineContent As String = CStr(currentWS.Cells(i, kapaConfig("LastLine").column).value)
                                                If Not IsNothing(lastLineContent) Then
                                                    Dim match As Match = regexpression.Match(lastLineContent)
                                                    If match.Success Then
                                                        lastLineContent = match.Value
                                                        found = True
                                                    End If
                                                End If
                                            End If

                                        End While
                                        lastZeile = i - 1


                                        ' letzte Zeile bestimmen, wenn dies verbunden Zellen sind
                                        ' -------------------------------------
                                        Dim rng As Range
                                        Dim rngEnd As Range

                                        rng = CType(currentWS.Cells(lastZeile, 1), Global.Microsoft.Office.Interop.Excel.Range)

                                        If rng.MergeCells Then

                                            rng = rng.MergeArea
                                            rngEnd = rng.Cells(rng.Rows.Count, rng.Columns.Count)

                                            ' dann ist die lastZeile neu zu besetzen
                                            lastZeile = rngEnd.Row
                                        End If
                                    End If


                                    If Not ok Then

                                        fehler = True

                                        If awinSettings.englishLanguage Then
                                            msgtxt = "Error reading availabilities: Please check the calendar in this file ..."
                                        Else
                                            msgtxt = "Fehler beim Lesen der Verfügbarkeiten: Bitte prüfen Sie die Korrektheit des Kalenders ..."
                                        End If
                                        If Not oPCollection.Contains(msgtxt) Then
                                            oPCollection.Add(msgtxt, msgtxt)
                                        End If
                                        'Call MsgBox(msgtxt)

                                        Call logger(ptErrLevel.logError, msgtxt, kapaFileName, anzFehler)

                                        If formerEE Then
                                            appInstance.EnableEvents = True
                                        End If

                                        If formerSU Then
                                            appInstance.ScreenUpdating = True
                                        End If

                                        enableOnUpdate = True
                                        If awinSettings.englishLanguage Then
                                            msgtxt = "Your availabilities couldn't be read, because of problems"
                                        Else
                                            msgtxt = "Ihre Verfügbarkeiten konnten nicht berücksichtigt werden"
                                        End If
                                        If Not oPCollection.Contains(msgtxt) Then
                                            oPCollection.Add(msgtxt, msgtxt)
                                        End If

                                        Call logger(ptErrLevel.logWarning, msgtxt, kapaFileName, anzFehler)
                                        'Call showOutPut(oPCollection, "Lesen Urlaubsplanung wurde mit Fehler abgeschlossen", "Meldungen zu Lesen Urlaubsplanung")
                                        ' tk 12.2.19 ess oll alles gelesen werden - es wird nicht weitergemacht, wenn es Einträge in der outputCollection gibt 
                                        'Throw New ArgumentException(msgtxt)
                                    Else

                                        For iZ = firstUrlzeile To lastZeile


                                            rolename = CType(currentWS.Cells(iZ, kapaConfig("role").column), Global.Microsoft.Office.Interop.Excel.Range).Text

                                            ' tk 31.1.2020 Test - der CheckWert steht auf Spalte "AS"
                                            ' dazu muss manuell der Check-Wert bestimmt und in der Excel Datei eingetragen werden ..  
                                            Dim checkWert As Double = -1
                                            Try
                                                If Not IsNothing(CType(currentWS.Cells(iZ, "AS"), Global.Microsoft.Office.Interop.Excel.Range).Value) Then
                                                    If IsNumeric(CType(currentWS.Cells(iZ, "AS"), Global.Microsoft.Office.Interop.Excel.Range).Value) Then
                                                        checkWert = CDbl(CType(currentWS.Cells(iZ, "AS"), Global.Microsoft.Office.Interop.Excel.Range).Value)
                                                    End If
                                                End If
                                            Catch ex As Exception
                                                checkWert = -1
                                            End Try
                                            ' Ende tk 31.1.2020 Auslesen Checkwert für Kapa-Bestimmung 

                                            If rolename <> "" Then
                                                hrole = RoleDefinitions.getRoledef(rolename)
                                                If Not IsNothing(hrole) Then

                                                    Dim defaultHrsPerdayForThisPerson As Double = hrole.defaultDayCapa

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

                                                                    Dim aktCell As Object = CType(currentWS.Cells(iZ, sp), Global.Microsoft.Office.Interop.Excel.Range).Value

                                                                    If Not IsNothing(CType(currentWS.Cells(iZ, sp), Global.Microsoft.Office.Interop.Excel.Range).Value) Then

                                                                        If IsNumeric(CType(currentWS.Cells(iZ, sp), Global.Microsoft.Office.Interop.Excel.Range).Value) Then

                                                                            Dim angabeInStd As Double = CType(CType(currentWS.Cells(iZ, sp), Global.Microsoft.Office.Interop.Excel.Range).Value, Double)

                                                                            If angabeInStd >= 0 And angabeInStd <= 24 Then
                                                                                anzArbStd = anzArbStd + CDbl(CType(currentWS.Cells(iZ, sp), Global.Microsoft.Office.Interop.Excel.Range).Value)
                                                                            Else
                                                                                If awinSettings.englishLanguage Then
                                                                                    msgtxt = "Error reading the amount of working hours for " & hrole.name & " : " & angabeInStd.ToString & " (!!)"
                                                                                Else
                                                                                    msgtxt = "Fehler beim Lesen der Anzahl zu leistenden Arbeitsstunden " & hrole.name & " : " & angabeInStd.ToString & " (!!)"
                                                                                End If
                                                                                If Not oPCollection.Contains(msgtxt) Then
                                                                                    oPCollection.Add(msgtxt, msgtxt)
                                                                                End If
                                                                                'Call MsgBox(msgtxt)
                                                                                fehler = True
                                                                                Call logger(ptErrLevel.logError, msgtxt, kapaFileName, anzFehler)
                                                                            End If
                                                                        Else
                                                                            Dim workHours As String = CType(currentWS.Cells(iZ, sp), Global.Microsoft.Office.Interop.Excel.Range).Value
                                                                            If workHours = "" Then
                                                                                ' Feld ist weiss, oder hat keine Farbe, keine Zahl und keinen "/": also ist es Arbeitstag mit Default-Std pro Tag 
                                                                                anzArbStd = anzArbStd + defaultHrsPerdayForThisPerson
                                                                            End If
                                                                            If kapaConfig("valueSign").regex = "RegEx" Then
                                                                                regexpression = New Regex(kapaConfig("valueSign").content)
                                                                                If Not IsNothing(workHours) Then
                                                                                    Dim match As Match = regexpression.Match(workHours)
                                                                                    If match.Success Then
                                                                                        workHours = match.Value
                                                                                        ' Feld ist weiss, oder hat keine Farbe, keine Zahl und keinen "/": also ist es Arbeitstag mit Default-Std pro Tag 
                                                                                        anzArbStd = anzArbStd + defaultHrsPerdayForThisPerson
                                                                                    End If
                                                                                End If
                                                                            End If

                                                                        End If

                                                                    Else
                                                                        ' ur:07.01.2020: Telair Variante entfällt mit Zeuss-Anpassung

                                                                        ' Feld ist ohne Inhalt: also ist es Arbeitstag mit Default-Std pro Tag 
                                                                        anzArbStd = anzArbStd + defaultHrsPerdayForThisPerson

                                                                        '' hier wird die Telair Variante gemacht 
                                                                        '' das einfachste wäre eigentlich  
                                                                        ''anzArbStd = anzArbStd + defaultHrsPerdayForThisPerson

                                                                        ''Dim colorIndup As Integer = CType(currentWS.Cells(iZ, sp), Global.Microsoft.Office.Interop.Excel.Range).Borders(XlBordersIndex.xlDiagonalUp).ColorIndex

                                                                        '' ' Wenn das Feld nicht durch einen Diagonalen Strich gekennzeichnet ist
                                                                        ''If CType(currentWS.Cells(iZ, sp), Global.Microsoft.Office.Interop.Excel.Range).Value <> "/" Then
                                                                        ''    'anzArbStd = anzArbStd + 8
                                                                        ''    anzArbStd = anzArbStd + defaultHrsPerdayForThisPerson
                                                                        ''Else
                                                                        ''    ' freier Tag für Teilzeitbeschäftigte
                                                                        ''    msgtxt = "Tag zählt nicht: Zeile " & iZ & ", Spalte " & sp
                                                                        ''    Call logfileSchreiben(msgtxt, kapaFileName, anzFehler)
                                                                        ''End If

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
                                                                Call logger(ptErrLevel.logError, msgtxt, kapaFileName, anzFehler)
                                                            End If

                                                        Next

                                                        anzArbTage = anzArbStd / 8

                                                        ' tk 31.1.20 Check den Wert
                                                        Dim formerVD As Boolean = awinSettings.visboDebug
                                                        awinSettings.visboDebug = True
                                                        If awinSettings.visboDebug Then
                                                            If checkWert <> -1 Then
                                                                If Math.Abs(anzArbTage - checkWert) > 0.0001 Then
                                                                    Call MsgBox("Abweichung in Kapa-Bestimmung")
                                                                End If
                                                            End If
                                                        End If
                                                        awinSettings.visboDebug = formerVD
                                                        'Ende tk Check den Wert 

                                                        'nur wenn die hrole schon eingetreten und nicht ausgetreten ist, wird die Capa eingetragen
                                                        If colOfDate >= getColumnOfDate(hrole.entryDate) And colOfDate < getColumnOfDate(hrole.exitDate) Then
                                                            hrole.kapazitaet(colOfDate) = anzArbTage
                                                        Else
                                                            hrole.kapazitaet(colOfDate) = 0
                                                        End If
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
                                                    Call logger(ptErrLevel.logError, msgtxt, kapaFileName, anzFehler)
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
                                                Call logger(ptErrLevel.logError, msgtxt, kapaFileName, anzFehler)
                                            End If

                                        Next iZ

                                    End If   ' ende von if not OK
                                Else

                                    If awinSettings.visboDebug Then

                                        If awinSettings.englishLanguage Then
                                            msgtxt = "Worksheet " & kapaFileName & "doesn't contain month/year ..."
                                        Else
                                            msgtxt = "Worksheet" & kapaFileName & " enthält keine Angaben zu Monat/Jahr ..."
                                        End If
                                        If Not oPCollection.Contains(msgtxt) Then
                                            oPCollection.Add(msgtxt, msgtxt)
                                        End If
                                        Call logger(ptErrLevel.logError, msgtxt, kapaFileName, anzFehler)
                                    End If

                                End If

                            End If      'beginningDay = 1

                        Next index


                    Catch ex2 As Exception
                        If awinSettings.englishLanguage Then
                            msgtxt = "Error reading dates like month/year ..."
                        Else
                            msgtxt = "Fehler beim Lesen der notwendigen Randdaten wie Monat/Jahr ..."
                        End If
                        If Not oPCollection.Contains(msgtxt) Then
                            oPCollection.Add(msgtxt, msgtxt)
                        End If
                        Call logger(ptErrLevel.logError, msgtxt, kapaFileName, anzFehler)
                    End Try

                    'kapaWB.Close(SaveChanges:=False)
                Catch ex As Exception

                End Try

            End If
        Else

        End If




        If formerEE Then
            appInstance.EnableEvents = True
        End If

        If formerSU Then
            appInstance.ScreenUpdating = True
        End If

        enableOnUpdate = True

        kapaWB.Close(SaveChanges:=False)

        readAvailabilityOfRoleWithConfig = (oPCollection.Count = old_oPCollectionCount)

    End Function

    ''' <summary>
    ''' 
    ''' </summary>
    ''' <param name="actualDataConfig"></param>
    ''' <param name="tmpDatei"></param>
    ''' <param name="oPCollection"></param>
    ''' <returns></returns>
    Public Function readCalendarReferenceFile(ByVal actualDataConfig As SortedList(Of String, clsConfigActualDataImport),
                                             ByVal tmpDatei As String,
                                             ByRef special As clsOtherCalendar,
                                             ByRef oPCollection As Collection) As Boolean

        Dim err As New clsErrorCodeMsg
        Dim outputline As String = ""
        Dim ok As Boolean = False
        Dim result As Boolean = True
        Dim actDataWB As Microsoft.Office.Interop.Excel.Workbook = Nothing
        Dim currentWS As Microsoft.Office.Interop.Excel.Worksheet = Nothing
        Dim searcharea As Microsoft.Office.Interop.Excel.Range = Nothing
        Dim t As Integer = 0  ' tabellenIndex
        Dim curmonth As Integer
        Dim stundenTotal As Integer = 0 ' Stundenangabe in einer Zeile
        Dim monat As String = ""
        Dim jahr As String = ""
        Dim yyyymm As String = ""

        ' ======================
        ' vorarbeit der Definitionen geleistet
        ' ======================
        Try
            If My.Computer.FileSystem.FileExists(tmpDatei) Then
                Try
                    ' Folgendes nur nötig, wenn die tmpDatei mit Signatur versehen ist
                    appInstance.DisplayAlerts = False
                    actDataWB = appInstance.Workbooks.Open(tmpDatei, UpdateLinks:=0)
                    actDataWB.Final = False
                    appInstance.DisplayAlerts = True

                    Dim vstart As clsConfigActualDataImport = actualDataConfig("valueStart")


                    ' Schleife über alle Tabellenblätter eines ausgewählten Excel-Files (hier = einer Rolle)
                    For t = 0 To vstart.sheet.bis - vstart.sheet.von

                        If Not IsNothing(vstart.sheet.von + t) Then
                            currentWS = CType(appInstance.Worksheets(vstart.sheet.von + t), Global.Microsoft.Office.Interop.Excel.Worksheet)
                            If Not IsNothing(vstart.sheetDescript) Then
                                ok = (vstart.sheetDescript.Contains(currentWS.Name))
                            Else
                                ok = True
                            End If
                        End If

                        If Not ok Then
                            If awinSettings.englishLanguage Then
                                outputline = "the sheet " & currentWS.Name & " doesn't match with the configuration"
                            Else
                                outputline = "das Tabellenblatt " & currentWS.Name & " passt nicht zur Konfiguration"
                            End If
                            oPCollection.Add(outputline)
                            Call logger(ptErrLevel.logError, outputline, "readActualDataWithConfig", anzFehler)
                            result = False
                            Exit For ' keine weiteren Tabellenblätter mehr lesen - Fehler aufgetreten
                        End If

                        If IsNothing(currentWS) Then
                            If awinSettings.englishLanguage Then
                                outputline = "the sheet " & vstart.sheetDescript & " doesn't exists in this workbook"
                            Else
                                outputline = "das Tabellenblatt " & vstart.sheetDescript & " ist nicht vorhanden"
                            End If
                            oPCollection.Add(outputline)
                            Call logger(ptErrLevel.logError, outputline, "readActualDataWithConfig", anzFehler)
                            result = False
                        Else
                            ' passendes Worksheet gefunden
                            Try
                                ' Find Month
                                monat = currentWS.Cells(actualDataConfig("months").row.von, actualDataConfig("months").column.von).value
                                Dim vglMonat As String = currentWS.Name
                                Dim validm As Boolean = (vglMonat.Contains(monat) Or monat.Contains(vglMonat))
                                ' find Year
                                jahr = currentWS.Cells(actualDataConfig("years").row.von, actualDataConfig("years").column.von).value
                                Dim vglJahr As String = currentWS.Name
                                Dim validj As Boolean = (vglJahr.Contains(jahr) Or jahr.Contains(vglJahr))

                                Dim xxx As Date = CDate("01." & monat & " " & jahr)
                                yyyymm = Format(xxx, "yyyy/MM")
                                curmonth = getColumnOfDate(xxx)

                            Catch ex As Exception
                                outputline = "Error looking for month/year"
                                oPCollection.Add(outputline)
                                Call logger(ptErrLevel.logError, outputline, "readActualDataWithConfig", anzFehler)
                                result = False
                            End Try

                            ' Find Wertespalte - auf jedem Tabellenblatt evt. anders
                            Dim hspalte As String = actualDataConfig("Total").columnDescript
                            Dim stdSpalteTotal As Integer = 0
                            Try
                                Dim überschriftenzeile As Integer = actualDataConfig("Überschriften").row.von
                                searcharea = currentWS.Rows(überschriftenzeile)          ' Zeile über... enthält die verschieden Spaltendescript
                                stdSpalteTotal = searcharea.Find(hspalte).Column
                                Dim filaWD As New clsFirstWDLastWD
                                Dim lastWorkDay As Date = CDate(currentWS.Cells(überschriftenzeile, stdSpalteTotal - 3).value.ToString & "." & monat & " " & jahr)
                                Dim hdate As Date = DateAdd(DateInterval.Month, -1, lastWorkDay)
                                Dim hmonth As String = MonthName(Month(hdate), Abbreviate:=True)
                                jahr = Year(hdate)
                                Dim firstWorkDay As Date = CDate(currentWS.Cells(überschriftenzeile, 6).value.ToString & "." & hmonth & " " & jahr)
                                filaWD.lastWorkDay = lastWorkDay
                                filaWD.firstWorkDay = firstWorkDay
                                special.otherCal.Add(yyyymm, filaWD)
                            Catch ex As Exception
                                If awinSettings.englishLanguage Then
                                    outputline = "Error: in the sheet " & vstart.sheetDescript & " the value-column " & hspalte & " not found"
                                Else
                                    outputline = "Error: im Tabellenblatt " & vstart.sheetDescript & " konnte die WerteSpalte " & hspalte & " nicht gefunden werden"
                                End If
                                oPCollection.Add(outputline)
                                Call logger(ptErrLevel.logError, outputline, "readActualDataWithConfig", anzFehler)
                                result = False
                            End Try

                        End If

                    Next t    ' nächste Tabelle des Excel-Inputfiles

                Catch ex As Exception
                    actDataWB = Nothing
                    Call MsgBox("1. " & ex.Message)
                End Try

                If Not IsNothing(actDataWB) Then
                    actDataWB.Close(SaveChanges:=False)
                End If


            End If
        Catch ex As Exception
            Call MsgBox("2. " & ex.Message)
        End Try


        readCalendarReferenceFile = result
    End Function


    ''' <summary>
    ''' liest das im Diretory ../ressource manager evt. liegende File 'zeuss*.xlsx' (oder wie in kapaConfig benamst) File  aus
    ''' und hinterlegt an entsprechender Stelle im hrole.kapazitaet die verfügbaren Tage der entsprechenden Rolle
    ''' </summary>
    ''' <remarks></remarks>
    Friend Function readAvailabilityOfRoleWithConfigCalendarReferenz(ByVal kapaConfig As SortedList(Of String, clsConfigKapaImport),
                                                                     ByVal calendarReference As clsOtherCalendar,
                                                                     ByVal referenzListe As SortedList(Of String, String),
                                                                     ByRef oPCollection As Collection) As Boolean

        Dim err As New clsErrorCodeMsg
        Dim old_oPCollectionCount As Integer = oPCollection.Count
        Dim relevantCapafiles As New SortedList(Of String, String)


        Dim ok As Boolean = True
        Dim formerEE As Boolean = appInstance.EnableEvents
        Dim formerSU As Boolean = appInstance.ScreenUpdating
        Dim msgtxt As String = ""
        Dim anzFehler As Integer = 0
        Dim fehler As Boolean = False

        Dim kapaWB As Microsoft.Office.Interop.Excel.Workbook = Nothing
        Dim myYear As Integer
        Dim myMonth As Integer
        Dim beginning As Date
        Dim ending As Date
        Dim spalte As Integer = 2
        Dim firstUrlspalte As Integer = 0
        Dim firstUrlzeile As Integer = 0
        Dim noColor As Integer = -4142
        Dim whiteColor As Integer = 2
        Dim currentWS As Excel.Worksheet
        Dim index As Integer
        Dim dateConsidered As Date

        'Dim year As Integer = DatePart(DateInterval.Year, Date.Now)
        Dim monthN As String = ""

        ' tk wird nicht verwendet ... 
        'Dim monthNumber As Integer = 0

        Dim Jahr As Integer = 0
        Dim anzMonthDays As Integer = 0
        Dim colOfDate As Integer = 0
        Dim anzDays As Integer = 0

        Dim lastZeile As Integer
        Dim lastSpalte As Integer
        Dim monthDays As New SortedList(Of Integer, Integer)

        Dim hrole As New clsRollenDefinition
        Dim rolename As String = ""

        Dim regexpression As Regex

        Dim outPutCollection As New Collection

        If formerEE Then
            appInstance.EnableEvents = False
        End If

        If formerSU Then
            appInstance.ScreenUpdating = False
        End If

        enableOnUpdate = False


        'Dim listOfFiles As Collections.ObjectModel.ReadOnlyCollection(Of String) = Nothing
        '' Dateien mit WildCards lesen
        'listOfFiles = My.Computer.FileSystem.GetFiles(importOrdnerNames(PTImpExp.Kapas),
        '                 FileIO.SearchOption.SearchTopLevelOnly, kapaConfig("Kapa-Datei").capacityFile)

        ' look for the first beginning and ending and then take the actualData
        For Each kvp As KeyValuePair(Of String, clsFirstWDLastWD) In calendarReference.otherCal

            relevantCapafiles = New SortedList(Of String, String)
            Dim relevantMonth As Date = CDate(kvp.Key)
            beginning = kvp.Value.firstWorkDay
            ending = kvp.Value.lastWorkDay
            ' search for the relevant inputfiles
            myMonth = Month(relevantMonth)
            myYear = Year(relevantMonth)
            'Dim filenameKapa As String = kapaConfig("Kapa-Datei").capacityFile
            'Dim hstr() As String = Split(filenameKapa, "*")
            'filenameKapa = hstr(hstr.Length - 1)
            Dim refKeyBeginn As String = Year(beginning).ToString & Month(beginning).ToString("D2")
            Dim refkeyEnd As String = Year(ending).ToString & Month(ending).ToString("D2")
            If referenzListe.ContainsKey(refKeyBeginn) And referenzListe.ContainsKey(refkeyEnd) Then

                relevantCapafiles.Add(Year(beginning).ToString & Month(beginning).ToString("D2"),
                                  My.Computer.FileSystem.CombinePath(importOrdnerNames(PTImpExp.Kapas),
                                  referenzListe(refKeyBeginn)))

                relevantCapafiles.Add(Year(ending).ToString & Month(ending).ToString("D2"),
                                  My.Computer.FileSystem.CombinePath(importOrdnerNames(PTImpExp.Kapas),
                                  referenzListe(refkeyEnd)))


                Dim isdate As Boolean = DateTime.TryParse(MonthName(myMonth) & " " & myYear.ToString, dateConsidered)

                'Dim beginningDay As Integer = -1
                'Dim endingDay As Integer = -1
                Dim erstertag = DateAndTime.Day(beginning)
                Dim letztertag = DateAndTime.Day(ending)

                If myYear <> 0 And MonthName(myMonth) <> "" Then

                    colOfDate = getColumnOfDate(dateConsidered)

                    monthDays.Clear()

                    anzMonthDays = DateTime.DaysInMonth(myYear, Month(beginning))
                    Dim anzDaysCapa As Long = DateDiff(DateInterval.Day, beginning, ending)
                    Dim anzDaysThisMonth As Long = anzMonthDays - erstertag + 1

                    If Not monthDays.ContainsKey(colOfDate) Then
                        monthDays.Add(colOfDate, anzDaysCapa)
                    End If

                    Dim existAllFiles As Boolean = True
                    Dim capaFiles() As String = Nothing
                    Dim notExistentCapaFiles As New Collection
                    ReDim capaFiles(relevantCapafiles.Count - 1)

                    Dim n As Integer = 0
                    Dim relMonth As String = ""

                    ' checking if all relevantCapafiles exist
                    For Each rCf As KeyValuePair(Of String, String) In relevantCapafiles
                        If My.Computer.FileSystem.FileExists(rCf.Value) Then
                            capaFiles(n) = rCf.Value
                            n = n + 1
                        Else
                            relMonth = myYear.ToString("D4") & myMonth.ToString("D2")
                            'notExistentCapaFiles.Add(relMonth, rCf.Value)
                            notExistentCapaFiles.Add(kvp.Key, rCf.Value)
                            existAllFiles = False
                        End If
                    Next



                    ' walking through the relevantCapafiles for capacities of the employee
                    If existAllFiles Then
                        ' nimmt die Rollennamen auf, die bereits in einem Teil der kapas berücksichtigt wurden
                        Dim listOfRolesN1 As New SortedList(Of String, String)

                        For n = 0 To capaFiles.Length - 1

                            Dim capaFile As String = capaFiles(n)

                            Try
                                kapaWB = appInstance.Workbooks.Open(capaFile)

                                Try
                                    For index = 1 To appInstance.Worksheets.Count

                                        currentWS = CType(appInstance.Worksheets(index), Global.Microsoft.Office.Interop.Excel.Worksheet)
                                        With currentWS

                                            'Dim regex As String = kapaConfig("month").regex
                                            'Dim Inhalt As String = kapaConfig("month").content

                                            ' Auslesen der Jahreszahl, falls vorhanden
                                            Dim hjahr As String = CStr(.Cells(kapaConfig("year").row, kapaConfig("year").column).value)
                                            If IsNothing(hjahr) Then
                                                Jahr = 0
                                            Else
                                                If kapaConfig("year").regex = "RegEx" Then
                                                    'regexpression = New Regex("[0-9]{4}")
                                                    regexpression = New Regex(kapaConfig("year").content)
                                                    Dim match As Match = regexpression.Match(hjahr)
                                                    If match.Success Then
                                                        Jahr = CInt(match.Value)
                                                        If myYear = Year(hjahr) Or myYear = Year(hjahr) - 1 Then
                                                            Jahr = myYear
                                                        End If
                                                    Else
                                                        Jahr = 0
                                                    End If
                                                End If
                                            End If

                                            ' Auslesen des relevanten Monats
                                            Dim hmonth As String = MonthName(myMonth)
                                            'Dim hmonth As String = CStr(.Cells(kapaConfig("month").row, kapaConfig("month").column).value)
                                            If IsNothing(hmonth) Then
                                                monthN = ""
                                            Else
                                                If kapaConfig("month").regex = "RegEx" Then
                                                    regexpression = New Regex(kapaConfig("month").content)
                                                    Dim Match As Match = regexpression.Match(hmonth)
                                                    If Match.Success Then
                                                        monthN = Match.Value
                                                        If monthN <> hmonth Then
                                                            monthN = hmonth
                                                        End If
                                                    Else
                                                        monthN = hmonth
                                                    End If
                                                End If
                                            End If


                                            ' Auslesen erste Verfügbarkeitsspalte
                                            firstUrlspalte = kapaConfig("valueStart").column
                                            firstUrlzeile = kapaConfig("valueStart").row
                                        End With

                                        ' hier ist sichergestellt, dass die erste Spalte mit 1 beginnt, die letzte Spalte dem Tag entspricht, mit dem der Monat endet
                                        If Jahr = 0 Or monthN = "" Then

                                            If awinSettings.visboDebug Then

                                                If awinSettings.englishLanguage Then
                                                    msgtxt = "Worksheet " & capaFile & "doesn't contain month/year ..."
                                                Else
                                                    msgtxt = "Worksheet" & capaFile & " enthält keine Angaben zu Monat/Jahr ..."
                                                End If
                                                If Not oPCollection.Contains(msgtxt) Then
                                                    oPCollection.Add(msgtxt, msgtxt)
                                                End If
                                                Call logger(ptErrLevel.logError, msgtxt, capaFile, anzFehler)
                                            End If
                                        Else
                                            ok = True
                                            anzDays = 0

                                            lastSpalte = CType(currentWS.Cells(firstUrlzeile - 1, 2000), Global.Microsoft.Office.Interop.Excel.Range).End(Excel.XlDirection.xlToLeft).Column
                                            lastZeile = CType(currentWS.Cells(2000, 1), Global.Microsoft.Office.Interop.Excel.Range).End(Excel.XlDirection.xlUp).Row

                                            ' Nachkorrektur gemäss Angabe in KonfigDate 'LastLine'
                                            Dim found As Boolean = False
                                            Dim i As Integer = lastZeile + 1
                                            While Not found
                                                i = i - 1
                                                If kapaConfig("LastLine").regex = "RegEx" Then
                                                    regexpression = New Regex(kapaConfig("LastLine").content)
                                                    Dim lastLineContent As String = CStr(currentWS.Cells(i, kapaConfig("LastLine").column).value)
                                                    If Not IsNothing(lastLineContent) Then
                                                        Dim match As Match = regexpression.Match(lastLineContent)
                                                        If match.Success Then
                                                            lastLineContent = match.Value
                                                            found = True
                                                        End If
                                                    End If
                                                End If

                                            End While
                                            lastZeile = i - 1


                                            ' letzte Zeile bestimmen, wenn dies verbunden Zellen sind
                                            ' -------------------------------------
                                            Dim rng As Range
                                            Dim rngEnd As Range

                                            rng = CType(currentWS.Cells(lastZeile, 1), Global.Microsoft.Office.Interop.Excel.Range)

                                            If rng.MergeCells Then

                                                rng = rng.MergeArea
                                                rngEnd = rng.Cells(rng.Rows.Count, rng.Columns.Count)

                                                ' dann ist die lastZeile neu zu besetzen
                                                lastZeile = rngEnd.Row
                                            End If



                                            For iZ = firstUrlzeile To lastZeile

                                                rolename = CType(currentWS.Cells(iZ, kapaConfig("role").column), Global.Microsoft.Office.Interop.Excel.Range).Text

                                                ' tk 31.1.2020 Test - der CheckWert steht auf Spalte "AS"
                                                ' dazu muss manuell der Check-Wert bestimmt und in der Excel Datei eingetragen werden ..  
                                                Dim checkWert As Double = -1
                                                Try
                                                    If Not IsNothing(CType(currentWS.Cells(iZ, "AS"), Global.Microsoft.Office.Interop.Excel.Range).Value) Then
                                                        If IsNumeric(CType(currentWS.Cells(iZ, "AS"), Global.Microsoft.Office.Interop.Excel.Range).Value) Then
                                                            checkWert = CDbl(CType(currentWS.Cells(iZ, "AS"), Global.Microsoft.Office.Interop.Excel.Range).Value)
                                                        End If
                                                    End If
                                                Catch ex As Exception
                                                    checkWert = -1
                                                End Try
                                                ' Ende tk 31.1.2020 Auslesen Checkwert für Kapa-Bestimmung 

                                                If rolename <> "" Then
                                                    hrole = RoleDefinitions.getRoledef(rolename)

                                                    If Not IsNothing(hrole) Then
                                                        ' Liste aufbauen für Rollen, deren Stunden bereits für den 1.Teil eingelesen wurde
                                                        If n = 0 Then
                                                            listOfRolesN1.Add(rolename, rolename)
                                                        End If

                                                        Dim defaultHrsPerdayForThisPerson As Double = hrole.defaultDayCapa

                                                        Dim anzDaysNow As Integer
                                                        Dim iSp As Integer
                                                        Dim anzArbTage As Double = 0
                                                        Dim anzArbStd As Double = 0

                                                        ' Start und Ende der Spalten-Auslesung bestimmen
                                                        If n = 0 Then
                                                            iSp = firstUrlspalte + erstertag - 1
                                                            anzDaysNow = anzMonthDays - erstertag + 1
                                                        End If
                                                        If n = 1 Then
                                                            iSp = firstUrlspalte
                                                            anzDaysNow = anzDaysCapa - anzDaysThisMonth + 1
                                                        End If


                                                        For sp = iSp + 0 To iSp + anzDaysNow - 1


                                                            If iSp <= lastSpalte Then

                                                                Dim hint As Integer = CInt(CType(currentWS.Cells(iZ, sp), Global.Microsoft.Office.Interop.Excel.Range).Interior.ColorIndex)

                                                                If CInt(CType(currentWS.Cells(iZ, sp), Global.Microsoft.Office.Interop.Excel.Range).Interior.ColorIndex) = noColor _
                                                                    Or CType(currentWS.Cells(iZ, sp), Global.Microsoft.Office.Interop.Excel.Range).Interior.ColorIndex = whiteColor Then

                                                                    Dim aktCell As Object = CType(currentWS.Cells(iZ, sp), Global.Microsoft.Office.Interop.Excel.Range).Value

                                                                    If Not IsNothing(CType(currentWS.Cells(iZ, sp), Global.Microsoft.Office.Interop.Excel.Range).Value) Then

                                                                        If IsNumeric(CType(currentWS.Cells(iZ, sp), Global.Microsoft.Office.Interop.Excel.Range).Value) Then

                                                                            Dim angabeInStd As Double = CType(CType(currentWS.Cells(iZ, sp), Global.Microsoft.Office.Interop.Excel.Range).Value, Double)

                                                                            If angabeInStd >= 0 And angabeInStd <= 24 Then
                                                                                anzArbStd = anzArbStd + CDbl(CType(currentWS.Cells(iZ, sp), Global.Microsoft.Office.Interop.Excel.Range).Value)
                                                                            Else
                                                                                If awinSettings.englishLanguage Then
                                                                                    msgtxt = "Error reading the amount of working hours for " & hrole.name & " : " & angabeInStd.ToString & " (!!)"
                                                                                Else
                                                                                    msgtxt = "Fehler beim Lesen der Anzahl zu leistenden Arbeitsstunden " & hrole.name & " : " & angabeInStd.ToString & " (!!)"
                                                                                End If
                                                                                If Not oPCollection.Contains(msgtxt) Then
                                                                                    oPCollection.Add(msgtxt, msgtxt)
                                                                                End If
                                                                                'Call MsgBox(msgtxt)
                                                                                fehler = True
                                                                                Call logger(ptErrLevel.logError, msgtxt, capaFile, anzFehler)
                                                                            End If
                                                                        Else
                                                                            Dim workHours As String = CType(currentWS.Cells(iZ, sp), Global.Microsoft.Office.Interop.Excel.Range).Value
                                                                            If workHours = "" Then
                                                                                ' Feld ist weiss, oder hat keine Farbe, keine Zahl und keinen "/": also ist es Arbeitstag mit Default-Std pro Tag 
                                                                                anzArbStd = anzArbStd + defaultHrsPerdayForThisPerson
                                                                            End If
                                                                            If kapaConfig("valueSign").regex = "RegEx" Then
                                                                                regexpression = New Regex(kapaConfig("valueSign").content)
                                                                                If Not IsNothing(workHours) Then
                                                                                    Dim match As Match = regexpression.Match(workHours)
                                                                                    If match.Success Then
                                                                                        workHours = match.Value
                                                                                        ' Feld ist weiss, oder hat keine Farbe, keine Zahl und keinen "/": also ist es Arbeitstag mit Default-Std pro Tag 
                                                                                        anzArbStd = anzArbStd + defaultHrsPerdayForThisPerson
                                                                                    End If
                                                                                End If
                                                                            End If

                                                                        End If

                                                                    Else
                                                                        ' ur:07.01.2020: Telair Variante entfällt mit Zeuss-Anpassung

                                                                        ' Feld ist ohne Inhalt: also ist es Arbeitstag mit Default-Std pro Tag 
                                                                        anzArbStd = anzArbStd + defaultHrsPerdayForThisPerson

                                                                        '' hier wird die Telair Variante gemacht 
                                                                        '' das einfachste wäre eigentlich  
                                                                        ''anzArbStd = anzArbStd + defaultHrsPerdayForThisPerson

                                                                        ''Dim colorIndup As Integer = CType(currentWS.Cells(iZ, sp), Global.Microsoft.Office.Interop.Excel.Range).Borders(XlBordersIndex.xlDiagonalUp).ColorIndex

                                                                        '' ' Wenn das Feld nicht durch einen Diagonalen Strich gekennzeichnet ist
                                                                        ''If CType(currentWS.Cells(iZ, sp), Global.Microsoft.Office.Interop.Excel.Range).Value <> "/" Then
                                                                        ''    'anzArbStd = anzArbStd + 8
                                                                        ''    anzArbStd = anzArbStd + defaultHrsPerdayForThisPerson
                                                                        ''Else
                                                                        ''    ' freier Tag für Teilzeitbeschäftigte
                                                                        ''    msgtxt = "Tag zählt nicht: Zeile " & iZ & ", Spalte " & sp
                                                                        ''    Call logfileSchreiben(msgtxt, kapaFileName, anzFehler)
                                                                        ''End If

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
                                                                Call logger(ptErrLevel.logError, msgtxt, capaFile, anzFehler)
                                                            End If

                                                        Next sp
                                                        ' nächste capafile öffnen
                                                        ' für alle mitarbeiter Std auslesen


                                                        anzArbTage = anzArbStd / 8

                                                        ' tk 31.1.20 Check den Wert
                                                        Dim formerVD As Boolean = awinSettings.visboDebug
                                                        awinSettings.visboDebug = True
                                                        If awinSettings.visboDebug Then
                                                            If checkWert <> -1 Then

                                                                If Math.Abs(anzArbTage - checkWert) > 0.0001 Then
                                                                    Call MsgBox("Abweichung in Kapa-Bestimmung")
                                                                End If
                                                            End If
                                                        End If
                                                        awinSettings.visboDebug = formerVD
                                                        'Ende tk Check den Wert 

                                                        ' erstes relevantCapafile
                                                        If n = 0 Then
                                                            'nur wenn die hrole schon eingetreten und nicht ausgetreten ist, wird die Capa eingetragen
                                                            If colOfDate >= getColumnOfDate(hrole.entryDate) And colOfDate < getColumnOfDate(hrole.exitDate) Then
                                                                hrole.kapazitaet(colOfDate) = anzArbTage
                                                            Else
                                                                hrole.kapazitaet(colOfDate) = 0
                                                            End If
                                                        End If

                                                        ' zweites relavantCapafile
                                                        If n = 1 Then
                                                            'nur wenn die hrole schon eingetreten und nicht ausgetreten ist, wird die Capa eingetragen
                                                            If colOfDate >= getColumnOfDate(hrole.entryDate) And colOfDate < getColumnOfDate(hrole.exitDate) Then
                                                                If listOfRolesN1.ContainsKey(hrole.name) Then
                                                                    ' aufaddieren nur, wenn die Rolle im ersten Teil auch schon berücksichtigt wurde
                                                                    hrole.kapazitaet(colOfDate) = hrole.kapazitaet(colOfDate) + anzArbTage
                                                                Else
                                                                    ' sonst Anzahl PT setzen
                                                                    hrole.kapazitaet(colOfDate) = anzArbTage
                                                                End If
                                                            Else
                                                                hrole.kapazitaet(colOfDate) = 0
                                                            End If
                                                        End If


                                                        iSp = iSp + anzDays
                                                        anzArbTage = 0              ' Anzahl Arbeitstage wieder zurücksetzen für den nächsten Monat
                                                        anzArbStd = 0               ' Anzahl zu leistender Arbeitsstunden wieder zurücksetzen für den nächsten Monat

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
                                                        Call logger(ptErrLevel.logError, msgtxt, capaFile, anzFehler)
                                                    End If

                                                End If

                                            Next iZ  ' loop Zeilen

                                        End If

                                    Next index

                                Catch ex2 As Exception
                                    If awinSettings.englishLanguage Then
                                        msgtxt = "Error reading dates like month/year ..."
                                    Else
                                        msgtxt = "Fehler beim Lesen der notwendigen Randdaten wie Monat/Jahr ..."
                                    End If
                                    If Not oPCollection.Contains(msgtxt) Then
                                        oPCollection.Add(msgtxt, msgtxt)
                                    End If
                                    Call logger(ptErrLevel.logError, msgtxt, capaFile, anzFehler)
                                End Try

                                kapaWB.Close(SaveChanges:=False)
                            Catch ex As Exception

                            End Try

                        Next    ' Capafiles

                        If awinSettings.englishLanguage Then
                            msgtxt = "Holidays of " & myYear & "/" & myMonth & " imported"
                        Else
                            msgtxt = "Für " & myYear & "/" & myMonth & " wurden Urlaubstage eingelesen"
                        End If

                        Call logger(ptErrLevel.logInfo, msgtxt, dateConsidered, anzFehler)
                    Else
                        If awinSettings.englishLanguage Then
                            msgtxt = "Holidays of " & myYear & "/" & myMonth & "not imported"
                        Else
                            msgtxt = "Für " & myYear & "/" & myMonth & " wurden keine Urlaubstage eingelesen" &
                                vbLf & "Datei existiert nicht: " & notExistentCapaFiles(kvp.Key)
                        End If
                        Call logger(ptErrLevel.logWarning, msgtxt, dateConsidered, anzFehler)
                    End If

                End If

            End If  ' end if referenzListe enthält die beiden Monate

        Next        ' calenderReference.otherCal

        If formerEE Then
            appInstance.EnableEvents = True
        End If

        If formerSU Then
            appInstance.ScreenUpdating = True
        End If

        enableOnUpdate = True


        readAvailabilityOfRoleWithConfigCalendarReferenz = (oPCollection.Count = old_oPCollectionCount)

    End Function

    ''' <summary>
    ''' Calculation of Feiertag beginning with Easter
    ''' </summary>
    ''' <param name="Datum"></param>
    ''' <returns></returns>
    Public Function officialHoliday(Datum As Date,
                                    Optional ByRef feiertagsliste As SortedList(Of Date, String) = Nothing) As String
        Dim J%, D%
        Dim O As Date
        J = Year(Datum)
        'Osterberechnung
        D = (((255 - 11 * (J Mod 19)) - 21) Mod 30) + 21
        Dim anzDaysBeginningFirstMarch As Integer = D + (D > 48) + 6 - ((J + J \ 4 + D + (D > 48) + 1) Mod 7)
        O = DateAdd(DateInterval.Day, anzDaysBeginningFirstMarch, DateSerial(J, 3, 1))
        'O = DateAdd(DateInterval.Day, D + (D > 48) + 6 -
        '((J + J \ 4 + D + (D > 48) + 1) Mod 7), DateSerial(J, 3, 1))

        Dim x As Date = DateSerial(J, 11, 18)
        Dim l As Long = DateDiff(DateInterval.Day, x, Date.MinValue)
        Dim y As Object = l Mod 7
        'Feiertage berechnen
        Select Case Datum
            Case DateSerial(J, 1, 1)
                officialHoliday = "Neujahr"
            Case DateSerial(J, 1, 6)
                officialHoliday = "Dreikönig*"
            Case DateAdd("D", -2, O)
                officialHoliday = "Karfreitag"
            Case O
                officialHoliday = "Ostersonntag"
            Case DateAdd("D", 1, O)
                officialHoliday = "Ostermontag"
            Case DateSerial(J, 5, 1)
                officialHoliday = "Erster Mai"
            Case DateAdd("D", 39, O)
                officialHoliday = "Christi Himmelfahrt"
            Case DateAdd("D", 49, O)
                officialHoliday = "Pfingstsonntag"
            Case DateAdd("D", 50, O)
                officialHoliday = "Pfingstmontag"
            Case DateAdd("D", 60, O)
                officialHoliday = "Fronleichnam*"
            Case DateSerial(J, 8, 15)
                officialHoliday = "Maria Himmelfahrt*"
            Case DateSerial(J, 10, 3)
                officialHoliday = "Deutsche Einheit"
            Case BussUndBettag(J)
                officialHoliday = "Buß- und Bettag*"
            Case DateSerial(J, 10, 31)
                officialHoliday = "Reformationstag*"
            Case DateSerial(J, 11, 1)
                officialHoliday = "Allerheiligen*"
            Case DateSerial(J, 12, 24)
                officialHoliday = "Heilig Abend*"
            Case DateSerial(J, 12, 25)
                officialHoliday = "EWeihnacht"
            Case DateSerial(J, 12, 26)
                officialHoliday = "ZWeihnacht"
            Case DateSerial(J, 12, 31)
                officialHoliday = "Silvester*"
            Case Else
                officialHoliday = ""
        End Select

        If Not IsNothing(feiertagsliste) Then
            'Dim feiertagsliste As New SortedList(Of Date, String)
            feiertagsliste.Add(DateSerial(J, 1, 1), "Neujahr")
            feiertagsliste.Add(DateSerial(J, 1, 6), "Dreikönig*")
            feiertagsliste.Add(DateAdd("D", -2, O), "Karfreitag")
            feiertagsliste.Add(O, "Ostersonntag")
            feiertagsliste.Add(DateAdd("D", 1, O), "Ostermontag")
            feiertagsliste.Add(DateSerial(J, 5, 1), "Erster Mai")
            feiertagsliste.Add(DateAdd("D", 39, O), "Christi Himmelfahrt")
            feiertagsliste.Add(DateAdd("D", 49, O), "Pfingstsonntag")
            feiertagsliste.Add(DateAdd("D", 50, O), "Pfingstmontag")
            feiertagsliste.Add(DateAdd("D", 60, O), "Fronleichnam*")
            feiertagsliste.Add(DateSerial(J, 8, 15), "Maria Himmelfahrt*")
            feiertagsliste.Add(DateSerial(J, 10, 3), "Deutsche Einheit")
            feiertagsliste.Add(BussUndBettag(J), "Buß- und Bettag*")
            feiertagsliste.Add(DateSerial(J, 10, 31), "Reformationstag*")
            feiertagsliste.Add(DateSerial(J, 11, 1), "Allerheiligen*")
            feiertagsliste.Add(DateSerial(J, 12, 24), "Heilig Abend*")
            feiertagsliste.Add(DateSerial(J, 12, 25), "EWeihnacht")
            feiertagsliste.Add(DateSerial(J, 12, 26), "ZWeihnacht")
            feiertagsliste.Add(DateSerial(J, 12, 31), "Silvester*")
        End If

    End Function

    Public Function BussUndBettag(ByVal Jahr As Long) As Date

        'Buss- und Bettag:
        'am Mittwoch vor dem letzten Sonntag im Kirchenjahr

        Dim t As Long
        Dim d As Date = Date.MinValue

        For t = 16 To 22
            d = DateSerial(Jahr, 11, t)
            If Weekday(d) = vbWednesday Then
                BussUndBettag = d
                Exit Function
            End If
        Next
        BussUndBettag = d
    End Function

    ''' <summary>
    ''' liest Projekte gemäß Konfiguration ein 
    ''' </summary>
    ''' <param name="listOfProjectFiles"></param>
    ''' <param name="projectConfig"></param>
    ''' <param name="meldungen"></param>
    ''' <returns></returns>
    Public Function readProjectsAllg(ByVal listOfProjectFiles As Collection,
                                     ByVal projectConfig As SortedList(Of String, clsConfigProjectsImport),
                                     ByRef meldungen As Collection) As List(Of String)

        Dim formerEE As Boolean = appInstance.EnableEvents
        Dim formerSU As Boolean = appInstance.ScreenUpdating
        Dim listOfArchivFiles As New List(Of String)
        Dim anzFehler As Integer = 0
        Dim result As Boolean = False

        If formerEE Then
            appInstance.EnableEvents = False
        End If

        If formerSU Then
            appInstance.ScreenUpdating = False
        End If

        enableOnUpdate = False



        If listOfProjectFiles.Count > 0 Then
            ' Öffnen des projectFile
            For Each tmpDatei As String In listOfProjectFiles
                Call logger(ptErrLevel.logInfo, "Einlesen Projekte " & tmpDatei, "", anzFehler)
                result = readProjectsWithConfig(projectConfig, tmpDatei, meldungen)

                If result Then
                    ' hier: merken der erfolgreich importierten Projects Dateien
                    listOfArchivFiles.Add(tmpDatei)
                End If
            Next

        Else
            Dim errMsg As String = "Es gibt keine Datei zur Projekt-Anlage" & vbLf _
                             & "Es wurden daher jetzt keine berücksichtigt"

            ' das sollte nicht dazu führen, dass nichts gemacht wird 
            'meldungen.Add(errMsg)
            'ur: 08.01.2020: endgültige meldung erst nachdem alle abgearbeitet wurden
            'Call MsgBox(errMsg)

            Call logger(ptErrLevel.logError, errMsg, "", anzFehler)
        End If

        If result Then
            readProjectsAllg = listOfArchivFiles
        Else
            readProjectsAllg = New List(Of String)
        End If

    End Function
    Function readProjectsWithConfig(ByVal projectConfig As SortedList(Of String, clsConfigProjectsImport),
                                    ByVal tmpDatei As String,
                                    ByRef meldungen As Collection) As Boolean
        Dim outputline As String = ""
        Dim ok As Boolean = False
        Dim result As Boolean = False
        Dim projectWB As Microsoft.Office.Interop.Excel.Workbook = Nothing
        Dim currentWS As Microsoft.Office.Interop.Excel.Worksheet = Nothing
        Dim regexpression As Regex
        Dim firstUrlspalte As Integer
        Dim firstUrlzeile As Integer
        Dim lastSpalte As Integer
        Dim lastZeile As Integer
        Dim anz_Proj_created As Integer = 0
        Dim anz_Proj_notCreated As Integer = 0

        ' Variables to create a Project
        Dim hproj As clsProjekt
        Dim pName As String = ""
        Dim vName As String = ""
        Dim vorlagenName As String = ""
        Dim startDate As Date
        Dim endDate As Date
        Dim budget As Double
        Dim sfit As Double
        Dim risk As Double
        Dim projectNummer As String = ""
        Dim description As String = ""
        Dim listOfCustomFields As New Collection
        Dim businessUnit As String = ""
        Dim responsible As String = ""
        Dim status As String = ""
        Dim zeile As Integer = 0
        Dim roleNames() As String = Nothing
        Dim roleValues() As Double = Nothing
        Dim roleListNameValues As New SortedList(Of String, Double())
        Dim costNames() As String = Nothing
        Dim costValues() As Double = Nothing
        Dim phNames() As String
        Dim przPhasenAnteile() As Double
        Dim combinedName As Boolean = True
        Dim createBudget As Boolean = True
        Dim createCostsRolesAnyhow As Boolean = True

        Dim monthVon As Integer = 0
        Dim monthBis As Integer = 0

        Dim noGo As Integer = 0   'Sobald diese Variable > 0 ist, wird das Projekt nicht importiert


        Try
            If My.Computer.FileSystem.FileExists(tmpDatei) Then

                Try

                    projectWB = appInstance.Workbooks.Open(tmpDatei)

                    Dim vstart As clsConfigProjectsImport = projectConfig("valueStart")
                    ' Auslesen erste Projekt-Spalte
                    firstUrlspalte = vstart.column.von
                    firstUrlzeile = vstart.row.von

                    If appInstance.Worksheets.Count > 0 Then

                        If Not IsNothing(vstart.sheet) Then
                            currentWS = CType(appInstance.Worksheets(vstart.sheet), Global.Microsoft.Office.Interop.Excel.Worksheet)
                            ok = (currentWS.Name = vstart.sheetDescript)
                        End If
                        If Not ok Then
                            If Not IsNothing(vstart.sheetDescript) Then
                                currentWS = CType(appInstance.Worksheets(vstart.sheetDescript), Global.Microsoft.Office.Interop.Excel.Worksheet)
                            Else
                                currentWS = Nothing
                            End If
                        End If

                        If IsNothing(currentWS) Then
                            outputline = "The Worksheet you want to import cannot be matched"
                            meldungen.Add(outputline)
                            Call logger(ptErrLevel.logError, outputline, "readProjectsWithConfig", anzFehler)
                        Else

                            lastSpalte = CType(currentWS.Cells(firstUrlzeile, 2000), Global.Microsoft.Office.Interop.Excel.Range).End(Excel.XlDirection.xlToLeft).Column
                            lastZeile = CType(currentWS.Cells(2000, firstUrlspalte), Global.Microsoft.Office.Interop.Excel.Range).End(Excel.XlDirection.xlUp).Row

                            Try
                                Dim projNumber As String = ""

                                For i = firstUrlzeile To lastZeile + 1

                                    'Find ProjectNumber
                                    Dim projNumber_new As String = ""
                                    Try
                                        Dim projNrConfig As clsConfigProjectsImport = projectConfig("ProjectNumber")

                                        If currentWS.Index <> projNrConfig.sheet Then
                                            If Not IsNothing(projNrConfig.sheet) Then
                                                currentWS = CType(appInstance.Worksheets(projNrConfig.sheet), Global.Microsoft.Office.Interop.Excel.Worksheet)
                                            Else
                                                currentWS = CType(appInstance.Worksheets(projNrConfig.sheetDescript), Global.Microsoft.Office.Interop.Excel.Worksheet)
                                            End If
                                        End If
                                        With currentWS

                                            Select Case projNrConfig.Typ
                                                Case "Text"
                                                    projNumber_new = CStr(.Cells(i, projNrConfig.column.von).value)
                                                Case "Integer"
                                                    projNumber_new = CInt(.Cells(i, projNrConfig.column.von).value)
                                                Case "Decimal"
                                                    projNumber_new = CDbl(.Cells(i, projNrConfig.column.von).value)
                                                Case "Date"
                                                    projNumber_new = CDate(.Cells(i, projNrConfig.column.von).value)
                                                Case Else
                                                    projNumber_new = .Cells(i, projNrConfig.column.von).value
                                            End Select

                                            If projNrConfig.objType = "RegEx" Then
                                                regexpression = New Regex(projNrConfig.content)
                                                Dim match As Match = regexpression.Match(projNumber_new)
                                                If match.Success Then
                                                    projNumber_new = match.Value
                                                Else
                                                    projNumber_new = Nothing
                                                End If
                                            End If

                                        End With

                                        If IsNothing(projNumber_new) Then
                                            If Not (i > lastZeile) Then
                                                If awinSettings.englishLanguage Then
                                                    outputline = "Couldn't find the projectnumber in line " & i.ToString & " of the inputfile"
                                                Else
                                                    outputline = "Fehler beim Herausfinden der Projektnummer in Zeile " & i.ToString & " des Inputfiles"
                                                End If
                                                'meldungen.Add(outputline)
                                                Call logger(ptErrLevel.logError, outputline, "readProjectsWithConfig", anzFehler)
                                                noGo = noGo + 1
                                                projNumber_new = projNumber
                                            End If
                                        Else

                                        End If

                                        If projNumber_new <> projNumber And i > firstUrlzeile Then
                                            If noGo > 0 Then
                                                If awinSettings.englishLanguage Then
                                                    outputline = "Error : Project '" & pName & "' starting at: " & startDate.ToString & " finishing at: " & endDate.ToString & "  N O T  imported !"
                                                Else
                                                    outputline = "Fehler : Projekt '" & pName & "' mit Start: " & startDate.ToString & " und Ende: " & endDate.ToString & "  N I C H T  erzeugt !"
                                                End If
                                                meldungen.Add(outputline)
                                                Call logger(ptErrLevel.logError, outputline, "readProjectsWithConfig", anzFehler)

                                                ' Zählen der aufgrund von fehlerhafter Definition o.ä. nicht erzeugten Projekten
                                                anz_Proj_notCreated = anz_Proj_notCreated + 1

                                                ' nach Projekt-Speicherung in ImportProjekte muss Bedarfsliste zurückgesetzt werden
                                                roleListNameValues = New SortedList(Of String, Double())

                                                ' zurücksetzen der Variable, die anzeigt, dass das aktuelle Projekt echte Fehler hatte beim Einlesen
                                                noGo = 0
                                            Else

                                                Dim anzRoles As Integer = roleListNameValues.Count
                                                ReDim roleNames(anzRoles - 1)
                                                ReDim roleValues(monthBis - monthVon)
                                                Dim k As Integer = 0
                                                For Each kvp As KeyValuePair(Of String, Double()) In roleListNameValues
                                                    roleNames(k) = kvp.Key
                                                    k = k + 1
                                                Next

                                                ReDim phNames(1)
                                                ReDim przPhasenAnteile(1)

                                                'erstelleProjektausParametern()
                                                anz_Proj_created = anz_Proj_created + 1
                                                hproj = New clsProjekt
                                                hproj = erstelleProjektausParametern(pName, vName, vorlagenName,
                                                                 startDate, endDate,
                                                                 budget, sfit, risk,
                                                                 projNumber, description,
                                                                 listOfCustomFields, businessUnit, responsible,
                                                                 status, zeile,
                                                                 roleNames, roleValues,
                                                                 costNames, costValues, phNames, przPhasenAnteile, combinedName, createBudget, createCostsRolesAnyhow)

                                                For Each kvp As KeyValuePair(Of String, Double()) In roleListNameValues

                                                    Dim tmpRCnameID As String = RoleDefinitions.bestimmeRoleNameID(kvp.Key, "")
                                                    hproj.AllPhases(0).getRoleByRoleNameID(tmpRCnameID).Xwerte = kvp.Value

                                                    Dim hilfe As Boolean = True
                                                Next

                                                ' Budget setzen 
                                                Call hproj.setBudgetAsNeeded()

                                                ' Beauftragen , weil aus Controlling Sheet kommt und Nummer hat 
                                                If hproj.kundenNummer <> "" Then
                                                    hproj.Status = ProjektStatus(PTProjektStati.beauftragt)
                                                End If


                                                ImportProjekte.Add(hproj, updateCurrentConstellation:=False)

                                                outputline = "Projekt '" & pName & "' mit Start: " & startDate.ToString & " und Ende: " & endDate.ToString & " erzeugt !"
                                                'meldungen.Add(outputline)
                                                Call logger(ptErrLevel.logInfo, outputline, "readProjectsWithConfig", anzFehler)

                                                ' nach Projekt-Speicherung in ImportProjekte muss Bedarfsliste zurückgesetzt werden
                                                roleListNameValues = New SortedList(Of String, Double())

                                            End If

                                        End If


                                        If i > lastZeile And IsNothing(projNumber_new) Then
                                            ' am Ende der zu lesenden Zeilen angekommen, die Felder sind nun leer
                                            ' beenden der Einlese-Aktion indem die For-Schleife abgebrochen wird
                                            Exit For
                                        End If

                                        projNumber = projNumber_new

                                    Catch ex As Exception
                                        If awinSettings.englishLanguage Then
                                            outputline = "Couldn't find the projectnumber in line " & i.ToString & "of the inputfile"
                                        Else
                                            outputline = "Fehler beim Herausfinden der ProjektNummer in Zeile " & i.ToString & " des Inputfiles"
                                        End If
                                        meldungen.Add(outputline)
                                        Call logger(ptErrLevel.logError, outputline, "readProjectsWithConfig", anzFehler)
                                        noGo = noGo + 1
                                    End Try

                                    'Find BusinesssUnit
                                    Dim projBU As Object
                                    Try
                                        Dim projBUConfig As clsConfigProjectsImport = projectConfig("BU")

                                        If currentWS.Index <> projBUConfig.sheet Then
                                            If Not IsNothing(projBUConfig.sheet) Then
                                                currentWS = CType(appInstance.Worksheets(projBUConfig.sheet), Global.Microsoft.Office.Interop.Excel.Worksheet)
                                            Else
                                                currentWS = CType(appInstance.Worksheets(projBUConfig.sheetDescript), Global.Microsoft.Office.Interop.Excel.Worksheet)
                                            End If
                                        End If
                                        With currentWS

                                            Select Case projBUConfig.Typ
                                                Case "Text"
                                                    projBU = CStr(.Cells(i, projBUConfig.column.von).value)
                                                Case "Integer"
                                                    projBU = CInt(.Cells(i, projBUConfig.column.von).value)
                                                Case "Decimal"
                                                    projBU = CDbl(.Cells(i, projBUConfig.column.von).value)
                                                Case "Color"
                                                    projBU = CLng(.Cells(i, projBUConfig.column.von).value)
                                                Case Else
                                                    projBU = .Cells(i, projBUConfig.column.von).value
                                            End Select

                                            If projBUConfig.objType = "RegEx" Then
                                                regexpression = New Regex(projBUConfig.content)
                                                Dim match As Match = regexpression.Match(projBU)
                                                If match.Success Then
                                                    projBU = match.Value
                                                End If
                                            End If
                                            businessUnit = projBU
                                        End With
                                    Catch ex As Exception
                                        If awinSettings.englishLanguage Then
                                            outputline = "Couldn't find the BU in line " & i.ToString & "of the inputfile"
                                        Else
                                            outputline = "Fehler beim Herausfinden der BU in Zeile " & i.ToString & " des Inputfiles"
                                        End If
                                        meldungen.Add(outputline)
                                        Call logger(ptErrLevel.logError, outputline, "readProjectsWithConfig", anzFehler)
                                        noGo = noGo + 1
                                    End Try


                                    'Find ProjectName
                                    Dim projName As String
                                    Try
                                        Dim projNameConfig As clsConfigProjectsImport = projectConfig("ProjectName")
                                        If currentWS.Index <> projNameConfig.sheet Then
                                            If Not IsNothing(projNameConfig.sheet) Then
                                                currentWS = CType(appInstance.Worksheets(projNameConfig.sheet), Global.Microsoft.Office.Interop.Excel.Worksheet)
                                            Else
                                                currentWS = CType(appInstance.Worksheets(projNameConfig.sheetDescript), Global.Microsoft.Office.Interop.Excel.Worksheet)
                                            End If
                                        End If

                                        With currentWS
                                            Select Case projNameConfig.Typ
                                                Case "Text"
                                                    projName = CStr(.Cells(i, projNameConfig.column.von).value)
                                                Case "Integer"
                                                    projName = CInt(.Cells(i, projNameConfig.column.von).value)
                                                Case "Decimal"
                                                    projName = CDbl(.Cells(i, projNameConfig.column.von).value)
                                                Case "Date"
                                                    projName = CDate(.Cells(i, projNameConfig.column.von).value)
                                                Case Else
                                                    projName = .Cells(i, projNameConfig.column.von).value
                                            End Select

                                            If projNameConfig.objType = "RegEx" Then
                                                regexpression = New Regex(projNameConfig.content)
                                                Dim match As Match = regexpression.Match(projName)
                                                If match.Success Then
                                                    projName = match.Value
                                                Else
                                                    projName = Nothing
                                                End If
                                            End If
                                            pName = projName
                                            ' ggfs. vorhandene Sonderzeichen wie (,),# [,] ersetzen
                                            If Not isValidPVName(pName) Then
                                                pName = makeValidProjectName(pName)
                                            End If

                                        End With
                                    Catch ex As Exception
                                        If awinSettings.englishLanguage Then
                                            outputline = "Couldn't find the projectname in line " & i.ToString & " of the inputfile"
                                        Else
                                            outputline = "Fehler beim Herausfinden des ProjektNamens in Zeile " & i.ToString & " des Inputfiles"
                                        End If
                                        'meldungen.Add(outputline)
                                        Call logger(ptErrLevel.logError, outputline, "readProjectsWithConfig", anzFehler)
                                        noGo = noGo + 1
                                    End Try


                                    'Find ProjectTemplate
                                    Dim projTmp As String
                                    Try
                                        Dim projTmpConfig As clsConfigProjectsImport = projectConfig("ProjectTemplate")
                                        If currentWS.Index <> projTmpConfig.sheet And projTmpConfig.sheet <> 0 Then
                                            If Not IsNothing(projTmpConfig.sheet) Then
                                                currentWS = CType(appInstance.Worksheets(projTmpConfig.sheet), Global.Microsoft.Office.Interop.Excel.Worksheet)
                                            Else
                                                currentWS = CType(appInstance.Worksheets(projTmpConfig.sheetDescript), Global.Microsoft.Office.Interop.Excel.Worksheet)
                                            End If
                                        End If

                                        If projTmpConfig.objType = "direkt" Then
                                            vorlagenName = projTmpConfig.content
                                        Else
                                            With currentWS
                                                Select Case projTmpConfig.Typ
                                                    Case "Text"
                                                        projTmp = CStr(.Cells(i, projTmpConfig.column.von).value)
                                                    Case "Integer"
                                                        projTmp = CInt(.Cells(i, projTmpConfig.column.von).value)
                                                    Case "Decimal"
                                                        projTmp = CDbl(.Cells(i, projTmpConfig.column.von).value)
                                                    Case "Date"
                                                        projTmp = CDate(.Cells(i, projTmpConfig.column.von).value)
                                                    Case Else
                                                        projTmp = .Cells(i, projTmpConfig.column.von).value
                                                End Select

                                                If projTmpConfig.objType = "RegEx" Then
                                                    regexpression = New Regex(projTmpConfig.content)
                                                    Dim match As Match = regexpression.Match(projTmp)
                                                    If match.Success Then
                                                        projTmp = match.Value
                                                    End If
                                                End If
                                                vorlagenName = projTmp
                                            End With
                                        End If
                                    Catch ex As Exception
                                        If awinSettings.englishLanguage Then
                                            outputline = "Couldn't find the project-template in line " & i.ToString & "of the inputfile"
                                        Else
                                            outputline = "Fehler beim Herausfinden des Projekt-Template in Zeile " & i.ToString & " des Inputfiles"
                                        End If
                                        meldungen.Add(outputline)
                                        Call logger(ptErrLevel.logError, outputline, "readProjectsWithConfig", anzFehler)
                                        noGo = noGo + 1
                                    End Try

                                    'Find ProjectStart
                                    Dim projStart As String = ""
                                    Try
                                        Dim projStartConfig As clsConfigProjectsImport = projectConfig("ProjectStart")
                                        If currentWS.Index <> projStartConfig.sheet And projStartConfig.sheet <> 0 Then
                                            If Not IsNothing(projStartConfig.sheet) Then
                                                currentWS = CType(appInstance.Worksheets(projStartConfig.sheet), Global.Microsoft.Office.Interop.Excel.Worksheet)
                                            Else
                                                currentWS = CType(appInstance.Worksheets(projStartConfig.sheetDescript), Global.Microsoft.Office.Interop.Excel.Worksheet)
                                            End If
                                        End If

                                        If projStartConfig.objType = "direkt" Then
                                            startDate = CDate(projStartConfig.content)
                                        Else
                                            With currentWS

                                                Select Case projStartConfig.Typ
                                                    Case "Text"
                                                        projStart = CStr(.Cells(i, projStartConfig.column.von).value)
                                                    Case "Integer"
                                                        projStart = CInt(.Cells(i, projStartConfig.column.von).value)
                                                    Case "Decimal"
                                                        projStart = CDbl(.Cells(i, projStartConfig.column.von).value)
                                                    Case "Date"
                                                        projStart = CDate(currentWS.Cells(i, projStartConfig.column.von).value)
                                                    Case Else
                                                        projStart = .Cells(i, projStartConfig.column.von).value
                                                End Select

                                                If projStartConfig.objType = "RegEx" Then
                                                    regexpression = New Regex(projStartConfig.content)
                                                    Dim match As Match = regexpression.Match(projStart)
                                                    If match.Success Then
                                                        projStart = match.Value
                                                    End If
                                                End If
                                            End With
                                            startDate = projStart
                                        End If


                                    Catch ex As Exception
                                        If awinSettings.englishLanguage Then
                                            outputline = "Couldn't find the projectstart in line " & i.ToString & "of the inputfile"
                                        Else
                                            outputline = "Fehler beim Herausfinden des Projekt-Starts in Zeile " & i.ToString & " des Inputfiles"
                                        End If
                                        meldungen.Add(outputline)
                                        Call logger(ptErrLevel.logError, outputline, "readProjectsWithConfig", anzFehler)
                                        noGo = noGo + 1
                                    End Try

                                    'Find ProjectEnde
                                    Dim projEnde As String = ""
                                    Try
                                        Dim projEndeConfig As clsConfigProjectsImport = projectConfig("ProjectEnd")
                                        If currentWS.Index <> projEndeConfig.sheet And projEndeConfig.sheet <> 0 Then
                                            If Not IsNothing(projEndeConfig.sheet) Then
                                                currentWS = CType(appInstance.Worksheets(projEndeConfig.sheet), Global.Microsoft.Office.Interop.Excel.Worksheet)
                                            Else
                                                currentWS = CType(appInstance.Worksheets(projEndeConfig.sheetDescript), Global.Microsoft.Office.Interop.Excel.Worksheet)
                                            End If
                                        End If

                                        If projEndeConfig.objType = "direkt" Then
                                            endDate = CDate(projEndeConfig.content)
                                        Else

                                            With currentWS
                                                Select Case projEndeConfig.Typ
                                                    Case "Text"
                                                        projEnde = CStr(.Cells(i, projEndeConfig.column.von).value)
                                                    Case "Integer"
                                                        projEnde = CInt(.Cells(i, projEndeConfig.column.von).value)
                                                    Case "Decimal"
                                                        projEnde = CDbl(.Cells(i, projEndeConfig.column.von).value)
                                                    Case "Date"
                                                        projEnde = CDate(.Cells(i, projEndeConfig.column.von).value)
                                                    Case Else
                                                        projEnde = .Cells(i, projEndeConfig.column.von).value
                                                End Select

                                                If projEndeConfig.objType = "RegEx" Then
                                                    regexpression = New Regex(projEndeConfig.content)
                                                    Dim match As Match = regexpression.Match(projEnde)
                                                    If match.Success Then
                                                        projEnde = match.Value
                                                    End If
                                                End If
                                            End With
                                            endDate = CDate(projEnde)
                                        End If
                                    Catch ex As Exception
                                        If awinSettings.englishLanguage Then
                                            outputline = "Couldn't find the projectend in line " & i.ToString & "of the inputfile"
                                        Else
                                            outputline = "Fehler beim Herausfinden des Projekt-Endes in Zeile " & i.ToString & " des Inputfiles"
                                        End If
                                        meldungen.Add(outputline)
                                        Call logger(ptErrLevel.logError, outputline, "readProjectsWithConfig", anzFehler)
                                        noGo = noGo + 1
                                    End Try

                                    ' find ProjectDescription
                                    Dim projDescr As String = ""
                                    Try
                                        Dim projDescrConfig As clsConfigProjectsImport = projectConfig("ProjectDescription")
                                        If currentWS.Index <> projDescrConfig.sheet And projDescrConfig.sheet <> 0 Then
                                            If Not IsNothing(projDescrConfig.sheet) Then
                                                currentWS = CType(appInstance.Worksheets(projDescrConfig.sheet), Global.Microsoft.Office.Interop.Excel.Worksheet)
                                            Else
                                                currentWS = CType(appInstance.Worksheets(projDescrConfig.sheetDescript), Global.Microsoft.Office.Interop.Excel.Worksheet)
                                            End If
                                        End If

                                        If projDescrConfig.objType = "direkt" Then
                                            description = CStr(projDescrConfig.content)
                                        Else

                                            With currentWS
                                                Select Case projDescrConfig.Typ
                                                    Case "Text"
                                                        projDescr = CStr(.Cells(i, projDescrConfig.column.von).value)
                                                    Case "Integer"
                                                        projDescr = CInt(.Cells(i, projDescrConfig.column.von).value)
                                                    Case "Decimal"
                                                        projDescr = CDbl(.Cells(i, projDescrConfig.column.von).value)
                                                    Case "Date"
                                                        projDescr = CDate(.Cells(i, projDescrConfig.column.von).value)
                                                    Case Else
                                                        projDescr = .Cells(i, projDescrConfig.column.von).value
                                                End Select

                                                If projDescrConfig.objType = "RegEx" Then
                                                    regexpression = New Regex(projDescrConfig.content)
                                                    Dim match As Match = regexpression.Match(projDescr)
                                                    If match.Success Then
                                                        projDescr = match.Value
                                                    End If
                                                End If
                                                description = projDescr
                                            End With
                                        End If
                                    Catch ex As Exception
                                        If awinSettings.englishLanguage Then
                                            outputline = "Couldn't find the projectdescription in line " & i.ToString & "of the inputfile"
                                        Else
                                            outputline = "Fehler beim Herausfinden der Projekt-Beschreibung in Zeile " & i.ToString & " des Inputfiles"
                                        End If
                                        meldungen.Add(outputline)
                                        Call logger(ptErrLevel.logError, outputline, "readProjectsWithConfig", anzFehler)
                                    End Try



                                    ' Find TimeUnit
                                    Dim timeUnit As String = ""
                                    Dim timeUnitConfig As clsConfigProjectsImport = projectConfig("TimeUnit")
                                    If currentWS.Index <> timeUnitConfig.sheet Then
                                        If Not IsNothing(timeUnitConfig.sheet) Then
                                            currentWS = CType(appInstance.Worksheets(timeUnitConfig.sheet), Global.Microsoft.Office.Interop.Excel.Worksheet)
                                        Else
                                            currentWS = CType(appInstance.Worksheets(timeUnitConfig.sheetDescript), Global.Microsoft.Office.Interop.Excel.Worksheet)
                                        End If
                                    End If

                                    With currentWS
                                        Select Case timeUnitConfig.Typ
                                            Case "Text"
                                                timeUnit = CStr(.Cells(i, timeUnitConfig.column.von).value)
                                            Case "Integer"
                                                timeUnit = CInt(.Cells(i, timeUnitConfig.column.von).value)
                                            Case "Decimal"
                                                timeUnit = CDbl(.Cells(i, timeUnitConfig.column.von).value)
                                            Case "Date"
                                                timeUnit = CDate(.Cells(i, timeUnitConfig.column.von).value)
                                            Case Else
                                                timeUnit = .Cells(i, timeUnitConfig.column.von).value
                                        End Select

                                        If timeUnitConfig.objType = "RegEx" Then
                                            regexpression = New Regex(timeUnitConfig.content)
                                            Dim timeUnitMatch As Match = regexpression.Match(timeUnit)
                                            Dim xx As MatchCollection = regexpression.Matches(timeUnit)
                                            If timeUnitMatch.Success Then
                                                timeUnit = CStr(timeUnitMatch.Value)
                                                ' find months
                                                Dim months As String = ""
                                                'Dim monthVon As Integer = 0
                                                'Dim monthBis As Integer = 0
                                                Dim monthsConfig As clsConfigProjectsImport = projectConfig("months")
                                                If currentWS.Index <> monthsConfig.sheet Then
                                                    If Not IsNothing(monthsConfig.sheet) Then
                                                        currentWS = CType(appInstance.Worksheets(monthsConfig.sheet), Global.Microsoft.Office.Interop.Excel.Worksheet)
                                                    Else
                                                        currentWS = CType(appInstance.Worksheets(monthsConfig.sheetDescript), Global.Microsoft.Office.Interop.Excel.Worksheet)
                                                    End If
                                                End If
                                                With currentWS
                                                    If Not IsNothing(monthsConfig.Typ) And
                                                        Not IsNothing(monthsConfig.column) Then

                                                        Select Case monthsConfig.Typ
                                                            Case "Text"
                                                                months = CStr(.Cells(i, monthsConfig.column.von).value)
                                                            Case "Integer"
                                                                months = CInt(.Cells(i, monthsConfig.column.von).value)
                                                            Case "Decimal"
                                                                months = CDbl(.Cells(i, monthsConfig.column.von).value)
                                                            Case "Date"
                                                                months = CDate(.Cells(i, monthsConfig.column.von).value)
                                                            Case Else
                                                                months = .Cells(i, monthsConfig.column.von).value
                                                        End Select
                                                    End If

                                                    If Not IsNothing(monthsConfig.cellrange) Then
                                                        If monthsConfig.cellrange Then
                                                            monthVon = monthsConfig.column.von
                                                            monthBis = monthsConfig.column.bis
                                                        End If
                                                    End If

                                                    If Not IsNothing(monthsConfig.objType) Then
                                                        If monthsConfig.objType = "RegEx" Then
                                                            regexpression = New Regex(monthsConfig.content)
                                                            Dim match As Match = regexpression.Match(months)
                                                            If match.Success Then
                                                                months = CInt(match.Value)
                                                            End If
                                                        End If
                                                    End If


                                                End With

                                                ' Find Role
                                                Dim roleName As String
                                                Dim roleNameConfig As clsConfigProjectsImport = projectConfig("Ressourcen")
                                                If currentWS.Index <> roleNameConfig.sheet Then
                                                    If Not IsNothing(roleNameConfig.sheet) Then
                                                        currentWS = CType(appInstance.Worksheets(roleNameConfig.sheet), Global.Microsoft.Office.Interop.Excel.Worksheet)
                                                    Else
                                                        currentWS = CType(appInstance.Worksheets(roleNameConfig.sheetDescript), Global.Microsoft.Office.Interop.Excel.Worksheet)
                                                    End If
                                                End If

                                                With currentWS
                                                    Select Case roleNameConfig.Typ
                                                        Case "Text"
                                                            roleName = CStr(.Cells(i, roleNameConfig.column.von).value)
                                                        Case "Integer"
                                                            roleName = CInt(.Cells(i, roleNameConfig.column.von).value)
                                                        Case "Decimal"
                                                            roleName = CDbl(.Cells(i, roleNameConfig.column.von).value)
                                                        Case "Date"
                                                            roleName = CDate(.Cells(i, roleNameConfig.column.von).value)
                                                        Case Else
                                                            roleName = .Cells(i, roleNameConfig.column.von).value
                                                    End Select

                                                    If roleNameConfig.objType = "RegEx" Then
                                                        If IsNothing(roleNameConfig.content) Then
                                                            'Fehlermeldung einbauen
                                                            If awinSettings.englishLanguage Then
                                                                outputline = "There is no regular expression defined in the config for getting the rolename"
                                                            Else
                                                                outputline = "Es wurde keine Regular Expression für die Ressource definiert"
                                                            End If
                                                            meldungen.Add(outputline)
                                                            Call logger(ptErrLevel.logError, outputline, "readProjectsWithConfig", anzFehler)
                                                            noGo = noGo + 1
                                                        Else
                                                            regexpression = New Regex(roleNameConfig.content)
                                                            Dim col As MatchCollection = regexpression.Matches(roleName)
                                                            ' Loop through Matches.
                                                            For Each m As Match In col
                                                                ' Access first Group and its value.
                                                                Dim g As Group = m.Groups(1)
                                                                roleName = g.Value
                                                            Next
                                                        End If

                                                    End If

                                                    If RoleDefinitions.containsName(roleName) Then
                                                        Dim hroleValues As Double()
                                                        ' initialisieren des Array
                                                        ReDim hroleValues(monthBis - monthVon)
                                                        For m = monthVon To monthBis
                                                            Select Case timeUnit
                                                                Case "hours", "hour"
                                                                    hroleValues(m - monthVon) = CDbl(.Cells(i, m).value) / 8
                                                                Case "days", "day"
                                                                    hroleValues(m - monthVon) = CDbl(.Cells(i, m).value)
                                                                Case "weeks", "week"
                                                                    hroleValues(m - monthVon) = CDbl(.Cells(i, m).value) * 5
                                                                Case "months", "month"
                                                                    hroleValues(m - monthVon) = CDbl(.Cells(i, m).value) * nrOfDaysMonth
                                                            End Select

                                                        Next
                                                        If Not roleListNameValues.ContainsKey(roleName) Then
                                                            ' liste aufbauen, die später dazu dient, das erstellte Projekt zu befüllen
                                                            roleListNameValues.Add(roleName, hroleValues)
                                                        Else
                                                            ' evt. aufsummieren der jeweiligen werte eines Monats
                                                        End If
                                                    Else
                                                        If awinSettings.englishLanguage Then
                                                            outputline = "Role " & roleName & " isn't defined in this VC"
                                                        Else
                                                            outputline = "Rolle " & roleName & " existiert in diesem VC nicht"
                                                        End If

                                                        'meldungen.Add(outputline)
                                                        Call logger(ptErrLevel.logInfo, outputline, "readProjectsWithConfig", anzFehler)

                                                    End If
                                                End With

                                            End If
                                        End If

                                    End With
                                Next i

                            Catch ex As Exception
                                If awinSettings.englishLanguage Then
                                    outputline = "The actual file isn't conform with the Configuration!"
                                Else
                                    outputline = "die ausgewählte Datei entspricht nicht der Konfiguration!"
                                End If

                                meldungen.Add(outputline)
                                Call logger(ptErrLevel.logError, outputline, "readProjectsWithConfig", anzFehler)
                                noGo = noGo + 1
                            End Try

                        End If

                    End If


                    projectWB.Close(SaveChanges:=False)

                Catch ex As Exception
                    If awinSettings.englishLanguage Then
                        outputline = "There is something wrong with the inputfile!"
                    Else
                        outputline = "Fehler im Inputfile!"
                    End If

                    meldungen.Add(outputline)
                    Call logger(ptErrLevel.logError, outputline, "readProjectsWithConfig", anzFehler)
                End Try
            Else
                If awinSettings.englishLanguage Then
                    outputline = "The file you selected doesn't exist!"
                Else
                    outputline = "Die ausgewählte Datei existiert nicht!"
                End If
                Call logger(ptErrLevel.logError, outputline, "readProjectsWithConfig", anzFehler)
            End If

        Catch ex As Exception

        End Try

        result = (anz_Proj_created = ImportProjekte.Count) And (anz_Proj_created > 0) And anz_Proj_notCreated <= 0


        If awinSettings.englishLanguage Then
            outputline = vbLf & anz_Proj_created.ToString & " projects created !"
            meldungen.Add(outputline)
            Call logger(ptErrLevel.logInfo, outputline, "readProjectsWithConfig", anzFehler)

            outputline = anz_Proj_notCreated & " projects were N O T  created !"
            meldungen.Add(outputline)
            Call logger(ptErrLevel.logInfo, outputline, "readProjectsWithConfig", anzFehler)
        Else
            outputline = vbLf & anz_Proj_created.ToString & " Projekte wurden erzeugt !"
            meldungen.Add(outputline)
            Call logger(ptErrLevel.logInfo, outputline, "readProjectsWithConfig", anzFehler)

            outputline = anz_Proj_notCreated & " Projekte wurden N I C H T  erzeugt !"
            meldungen.Add(outputline)
            Call logger(ptErrLevel.logInfo, outputline, "readProjectsWithConfig", anzFehler)
        End If


        readProjectsWithConfig = result
    End Function

    Public Function readProjectsJIRA(ByVal listOfProjectFiles As Collection,
                                     ByVal projectConfig As SortedList(Of String, clsConfigProjectsImport),
                                     ByRef meldungen As Collection) As List(Of String)

        Dim formerEE As Boolean = appInstance.EnableEvents
        Dim formerSU As Boolean = appInstance.ScreenUpdating
        Dim JiraProjLength As Integer = 2
        Dim listOfArchivFiles As New List(Of String)
        Dim projtaskList As New SortedList(Of String, SortedList(Of Date, clsJIRA_Task))
        Dim projListSortedName As New SortedList(Of String, SortedList(Of String, clsJIRA_Task))
        Dim taskListSortedID As New SortedList(Of String, clsJIRA_Task)
        Dim tasksInserted As New SortedList(Of String, clsJIRA_Task)
        Dim tasksRemaining As New SortedList(Of String, clsJIRA_Task)
        Dim tasksBacklog As New SortedList(Of String, clsJIRA_Task)
        Dim tasksFertigOSprint As New SortedList(Of String, clsJIRA_Task)

        Dim projsprintList As New SortedList(Of String, SortedList(Of String, clsJIRA_sprint))
        Dim anzFehler As Integer = 0
        Dim result As Boolean = False

        If formerEE Then
            appInstance.EnableEvents = False
        End If

        If formerSU Then
            appInstance.ScreenUpdating = False
        End If

        enableOnUpdate = False



        If listOfProjectFiles.Count > 0 Then
            ' Öffnen des projectFile
            For Each tmpDatei As String In listOfProjectFiles
                Call logger(ptErrLevel.logInfo, "Einlesen JIRA-Projekte " & tmpDatei, "readProjectsJIRA", anzFehler)

                result = readJIRATasks(projectConfig, tmpDatei, projtaskList, projListSortedName, projsprintList, meldungen)

                If Not result Then
                    If awinSettings.englishLanguage Then
                        Call showOutPut(meldungen, " Import Jira-Projects", "Errors that occured")
                    Else
                        Call showOutPut(meldungen, " Import Jira-Projekte", "folgende Fehler sind aufgetreten")
                    End If
                Else

                    Call logger(ptErrLevel.logInfo, "JIRA-Projekte eingelesen", "readProjectsJIRA", anzFehler)

                    For Each kvp As KeyValuePair(Of String, SortedList(Of String, clsJIRA_Task)) In projListSortedName

                        Dim projectName As String = kvp.Key
                        Dim taskList As SortedList(Of Date, clsJIRA_Task) = projtaskList(projectName)
                        Dim sprintList As SortedList(Of String, clsJIRA_sprint) = projsprintList(projectName)
                        Dim taskListSorted As SortedList(Of String, clsJIRA_Task) = kvp.Value

                        ' Neu Initialisierung bei jedem anderen Projekt
                        tasksInserted = New SortedList(Of String, clsJIRA_Task)
                        tasksRemaining = New SortedList(Of String, clsJIRA_Task)
                        tasksBacklog = New SortedList(Of String, clsJIRA_Task)
                        tasksFertigOSprint = New SortedList(Of String, clsJIRA_Task)

                        ' Bestimme SprintEndDate vom letzen Sprint 
                        Dim lastSprintEnd As Date = StartofCalendar
                        Dim firstSprintStart As Date = EndOfCalendar
                        For Each sprint As KeyValuePair(Of String, clsJIRA_sprint) In sprintList
                            If sprint.Value.SprintStartDate < firstSprintStart Then
                                firstSprintStart = sprint.Value.SprintStartDate
                            End If
                            If sprint.Value.SprintEndDate > lastSprintEnd Then
                                lastSprintEnd = sprint.Value.SprintEndDate
                            End If
                        Next



                        Dim projStart As Date = StartofCalendar

                        ' Projekt erzeugen mit Name, Start- und Ende-Datum
                        Dim hproj As clsProjekt
                        Dim anzTasks As Integer = taskList.Count

                        ' taskList nach Vorgangstyp filtern
                        Dim epicCollection As SortedList(Of Date, clsJIRA_Task) = filternNach("Vorgangstyp", "Epic", taskList)

                        ' find the project-Start
                        For Each item As KeyValuePair(Of Date, clsJIRA_Task) In epicCollection
                            If Not IsNothing(item.Value.SprintName) Then
                                If (item.Value.SprintStartDate > Date.MinValue) And (item.Value.SprintStartDate < projStart) Then
                                    projStart = item.Value.SprintStartDate
                                End If
                            Else
                                If (item.Value.StartDate > Date.MinValue) And (item.Value.StartDate < projStart) Then
                                    projStart = item.Value.StartDate
                                End If
                            End If

                        Next
                        If projStart <= StartofCalendar Or projStart >= EndOfCalendar Then

                            ' StartDatum fürs Projekt ist dann das Erstellt-Datum der ersten Task, die nicht nothing ist,  in Tasklist
                            ' Projektstart bestimmen
                            Dim i As Integer = 0
                            While taskList.ElementAt(i).Key <= Date.MinValue
                                i = i + 1
                            End While
                            projStart = taskList.ElementAt(i).Key.Date  ' Uhrzweit wird nicht berücksichtigt
                        End If

                        If epicCollection.Count > 0 Then
                            Dim oneEpic As clsJIRA_Task = epicCollection.ElementAt(0).Value
                        End If


                        ' find the project-End
                        'Dim projEnde As Date = EndOfCalendar
                        Dim projEnde As Date = projStart.AddYears(JiraProjLength)

                        For Each item As KeyValuePair(Of Date, clsJIRA_Task) In epicCollection
                            If item.Value.SprintEndDate > projEnde Then
                                projEnde = item.Value.SprintEndDate
                            End If
                            If item.Value.fällig > projEnde Then
                                projEnde = item.Value.fällig
                            End If
                            If item.Value.erledigt > projEnde Then
                                projEnde = item.Value.erledigt
                            End If
                            If item.Value.aktualisiert > projEnde Then
                                projEnde = item.Value.aktualisiert
                            End If
                        Next

                        hproj = New clsProjekt(projectName, False, projStart.Date, projEnde.Date)


                        ' jeden epic-Vorgang einzeln filtern und als Phase in das Projekt eintragen
                        Dim epics() As SortedList(Of Date, clsJIRA_Task)
                        ReDim epics(epicCollection.Count - 1)

                        Dim ie As Integer = 0

                        Try

                            ' Phasen aus den Epics erzeugen
                            Dim ephase As clsPhase

                            For Each item As KeyValuePair(Of Date, clsJIRA_Task) In epicCollection

                                ephase = New clsPhase(hproj)
                                Dim ephasenameNew As String = item.Value.Jira_ID & " " & item.Value.Zusammenfassung
                                ephase.nameID = calcHryElemKey(ephasenameNew, False)

                                ' falls Synonyme definiert sind, ersetzen durch Std-Name, sonst bleibt Name unverändert 
                                Dim origPhName As String = ephase.name
                                ephasenameNew = phaseMappings.mapToStdName(".", ephasenameNew)

                                ' nachsehen, ob newPhaseName in PhaseDefinitions definiert ist
                                If Not PhaseDefinitions.Contains(ephasenameNew) Then
                                    Dim newPhaseDef As New clsPhasenDefinition
                                    newPhaseDef.name = ephasenameNew
                                    ' Abbreviation, falls Customfield visbo_abbrev definiert ist
                                    'If visbo_abbrev <> 0 Then          ' VISBO-Abbrev ist definiert
                                    '    newPhaseDef.shortName = msTask.GetField(visbo_abbrev)
                                    'Else
                                    newPhaseDef.shortName = item.Value.Jira_ID
                                    'End If

                                    ' Task Class, falls Customfield visbo_taskclass definiert ist
                                    'If visbo_taskclass <> 0 Then          ' VISBO-TaskClass ist definiert
                                    '    newPhaseDef.darstellungsKlasse = msTask.GetField(visbo_taskclass)
                                    'Else
                                    newPhaseDef.darstellungsKlasse = ""
                                    'End If
                                    ephase.appearanceName = newPhaseDef.darstellungsKlasse

                                    newPhaseDef.UID = PhaseDefinitions.Count + 1
                                    'PhaseDefinitions.Add(newPhaseDef)
                                    missingPhaseDefinitions.Add(newPhaseDef)
                                Else
                                    'If visbo_taskclass <> 0 Then          ' VISBO-TaskClass ist definiert
                                    '    cphase.appearanceName = msTask.GetField(visbo_taskclass)
                                    'Else
                                    ephase.appearanceName = appearanceDefinitions.getPhaseAppearance(ephasenameNew, "").name
                                    'End If
                                End If

                                Dim duration As Integer
                                Dim ephaseStart As Date
                                Dim ephaseEnd As Date

                                ' Bestimme epic-Start und Epic-Ende

                                If Not IsNothing(item.Value.StartDate) And item.Value.StartDate > Date.MinValue Then
                                    If (item.Value.StartDate >= hproj.startDate) Then
                                        ephaseStart = item.Value.StartDate
                                    Else
                                        ephaseStart = hproj.startDate
                                    End If
                                    ephaseEnd = projEnde
                                    If Not IsNothing(item.Value.fällig) And item.Value.fällig > Date.MinValue Then
                                        ephaseEnd = item.Value.fällig
                                    End If

                                    If Not IsNothing(item.Value.erledigt) And item.Value.erledigt > Date.MinValue Then
                                        ephaseEnd = item.Value.erledigt
                                    End If
                                    duration = calcDauerIndays(ephaseStart, ephaseEnd)

                                Else
                                    If Not IsNothing(item.Value.SprintName) Then          ' Epic wurde einem Sprint zugeordnet
                                        ephaseStart = item.Value.SprintStartDate
                                        ephaseEnd = item.Value.SprintEndDate
                                        If item.Value.SprintCompleteDate > Date.MinValue And item.Value.SprintCompleteDate < Date.MaxValue Then
                                            ephaseEnd = item.Value.SprintCompleteDate
                                        End If
                                        duration = calcDauerIndays(ephaseStart, ephaseEnd)
                                    Else
                                        ephaseStart = hproj.startDate
                                        If Not IsNothing(item.Value.Erstellt) And item.Value.Erstellt > Date.MinValue Then
                                            ephaseStart = item.Value.Erstellt
                                        End If
                                        ' Phase kann nicht vor dem Projekt beginnen
                                        If ephaseStart < hproj.startDate Then
                                            ephaseStart = hproj.startDate
                                        End If
                                        'If Not IsNothing(item.Value.erledigt) Then
                                        '    ephaseEnd = item.Value.erledigt
                                        'End If
                                        'ephaseEnd = item.Value.fällig

                                        duration = calcDauerIndays(ephaseStart, ephaseEnd)

                                        If duration < 0 Then
                                            ephaseEnd = projEnde
                                            If Not IsNothing(item.Value.fällig) And item.Value.fällig > Date.MinValue Then
                                                ephaseEnd = item.Value.fällig
                                            End If

                                            If Not IsNothing(item.Value.erledigt) And item.Value.erledigt > Date.MinValue Then
                                                ephaseEnd = item.Value.erledigt
                                            End If
                                            If item.Value.fällig <= ephaseStart Then
                                                If ephaseStart <= item.Value.aktualisiert Then
                                                    ephaseEnd = item.Value.aktualisiert
                                                Else
                                                    ephaseEnd = projEnde
                                                End If
                                            End If
                                            If ephaseEnd <= Date.MinValue Then
                                                ephaseEnd = projEnde
                                            End If

                                            'Call MsgBox("Phase " & item.Jira_ID & " hat eine negative Dauer: " & item.Erstellt & " " & item.fällig)
                                            duration = calcDauerIndays(ephaseStart, ephaseEnd)
                                        End If

                                    End If
                                End If


                                If duration > 0 Then
                                    Dim offset As Integer = DateDiff(DateInterval.Day, hproj.startDate.Date, ephaseStart.Date)
                                    ephase.offset = offset
                                    ephase.changeStartandDauer(offset, duration)
                                    If item.Value.TaskStatus = "Fertig" Then
                                        ephase.percentDone = 1.0
                                    Else
                                        ephase.percentDone = 0.0
                                    End If
                                    ephase.verantwortlich = item.Value.zugewPerson

                                    ' hphase in Hierarchie auf Level 1 eintragen und in Projekt einhängen
                                    Dim hrchynode As New clsHierarchyNode
                                    hrchynode.elemName = ephase.name
                                    hrchynode.parentNodeKey = rootPhaseName
                                    hproj.AddPhase(ephase, origName:=origPhName, parentID:=rootPhaseName)
                                    hrchynode.indexOfElem = hproj.AllPhases.Count
                                    tasksInserted.Add(item.Value.Jira_ID, item.Value)
                                End If

                                '  Tasks filtern nach JIRA_ID des Epics
                                Dim epicStoryPoints As Double = 0.0
                                Dim epicVorg As New SortedList(Of Date, clsJIRA_Task)
                                epicVorg = filternNach("Übergeordnet", item.Value.Jira_ID, taskList)
                                epics(ie) = epicVorg
                                Dim vgphase As clsPhase
                                Dim phaseStart As Date = ephase.getStartDate
                                Dim phaseEnd As Date = ephase.getEndDate
                                Dim erledigteVgCount As Integer = 0

                                For Each itemVg As KeyValuePair(Of Date, clsJIRA_Task) In epicVorg

                                    vgphase = New clsPhase(hproj)
                                    Dim vgphaseNameNew As String = itemVg.Value.Jira_ID & " " & itemVg.Value.Zusammenfassung
                                    vgphase.nameID = calcHryElemKey(vgphaseNameNew, False)

                                    ' falls Synonyme definiert sind, ersetzen durch Std-Name, sonst bleibt Name unverändert 
                                    Dim origPhNameVG As String = vgphaseNameNew
                                    vgphaseNameNew = phaseMappings.mapToStdName("", vgphaseNameNew)


                                    ' nachsehen, ob msTask.Name in PhaseDefinitions definiert ist
                                    If Not PhaseDefinitions.Contains(vgphaseNameNew) Then
                                        Dim newPhaseDef As New clsPhasenDefinition
                                        newPhaseDef.name = vgphaseNameNew
                                        ' Abbreviation, falls Customfield visbo_abbrev definiert ist
                                        'If visbo_abbrev <> 0 Then          ' VISBO-Abbrev ist definiert
                                        '    newPhaseDef.shortName = msTask.GetField(visbo_abbrev)
                                        'Else
                                        newPhaseDef.shortName = itemVg.Value.Jira_ID
                                        'End If

                                        ' Task Class, falls Customfield visbo_taskclass definiert ist
                                        'If visbo_taskclass <> 0 Then          ' VISBO-TaskClass ist definiert
                                        '    newPhaseDef.darstellungsKlasse = msTask.GetField(visbo_taskclass)
                                        'Else
                                        newPhaseDef.darstellungsKlasse = ""
                                        'End If
                                        vgphase.appearanceName = newPhaseDef.darstellungsKlasse

                                        newPhaseDef.UID = PhaseDefinitions.Count + 1
                                        'PhaseDefinitions.Add(newPhaseDef)
                                        missingPhaseDefinitions.Add(newPhaseDef)
                                    Else
                                        'If visbo_taskclass <> 0 Then          ' VISBO-TaskClass ist definiert
                                        '    cphase.appearanceName = msTask.GetField(visbo_taskclass)
                                        'Else
                                        vgphase.appearanceName = appearanceDefinitions.getPhaseAppearance(vgphaseNameNew, "").name
                                        'End If
                                    End If


                                    ' Bestimmung von Start und Ende der Task - evt. durch Sprint definiert
                                    Dim durationVg As Integer = 0
                                    Dim taskStart As Date
                                    Dim taskEnd As Date

                                    If Not IsNothing(itemVg.Value.SprintName) Then          ' Task wurde einem Sprint zugeordnet
                                        taskStart = itemVg.Value.SprintStartDate
                                        taskEnd = itemVg.Value.SprintEndDate
                                        If itemVg.Value.SprintCompleteDate > Date.MinValue And itemVg.value.sprintCompleteDate < Date.MaxValue Then
                                            taskEnd = itemVg.Value.SprintCompleteDate
                                        End If
                                        durationVg = calcDauerIndays(taskStart, taskEnd)
                                    Else
                                        taskStart = ephaseStart
                                        If Not IsNothing(itemVg.Value.StartDate) And itemVg.Value.StartDate > Date.MinValue Then
                                            taskStart = itemVg.Value.StartDate
                                        End If
                                        If Not IsNothing(itemVg.Value.Erstellt) And itemVg.Value.Erstellt > Date.MinValue Then
                                            taskStart = itemVg.Value.Erstellt
                                        End If
                                        If Not IsNothing(itemVg.Value.aktualisiert) And itemVg.Value.aktualisiert > Date.MinValue Then
                                            taskStart = itemVg.Value.aktualisiert
                                        End If

                                        If Not IsNothing(itemVg.Value.fällig) And itemVg.Value.fällig > Date.MinValue Then
                                            taskEnd = item.Value.fällig
                                        End If
                                        If Not IsNothing(item.Value.erledigt) And item.Value.erledigt > Date.MinValue Then
                                            taskEnd = item.Value.erledigt
                                        End If

                                        durationVg = calcDauerIndays(taskStart, taskEnd)
                                        If durationVg < 0 Then

                                            If Not IsNothing(itemVg.Value.erledigt) Then
                                                taskEnd = itemVg.Value.erledigt
                                            End If
                                            If itemVg.Value.fällig <= taskStart Then
                                                If taskStart <= itemVg.Value.aktualisiert Then
                                                    taskEnd = itemVg.Value.aktualisiert
                                                Else
                                                    taskEnd = projEnde
                                                End If
                                            End If
                                            If taskEnd <= Date.MinValue Then
                                                taskEnd = ephaseEnd
                                            End If

                                            'Call MsgBox("Phase " & item.Jira_ID & " hat eine negative Dauer: " & item.Erstellt & " " & item.fällig)
                                            durationVg = calcDauerIndays(taskStart, itemVg.Value.fällig)
                                        End If

                                    End If

                                    ' Anfang der Task kann nicht vor dem Start des Epic sein
                                    If (taskStart < ephaseStart) Then
                                        taskStart = ephaseStart
                                    End If

                                    If durationVg > 0 Then
                                        Dim offset As Integer = DateDiff(DateInterval.Day, hproj.startDate.Date, taskStart.Date)
                                        vgphase.offset = offset
                                        vgphase.changeStartandDauer(offset, durationVg)
                                        If itemVg.Value.TaskStatus = "Fertig" Then
                                            vgphase.percentDone = 1.0
                                        Else
                                            vgphase.percentDone = 0.0
                                        End If
                                        vgphase.verantwortlich = itemVg.Value.zugewPerson


                                        ' Ressources auf Vorgang verteilen
                                        Dim vgXwerte As Double() = Nothing
                                        Dim vgoldXwerte As Double()
                                        Dim vganfang As Integer = vgphase.relStart
                                        Dim vgende As Integer = vgphase.relEnde
                                        If itemVg.Value.StoryPoints > 0.0 Then

                                            ' Aufsammeln der StoryPoints aller Tasks zu einer Epic
                                            epicStoryPoints = epicStoryPoints + itemVg.Value.StoryPoints

                                            ' ein StoryPoint in JIRA entspricht  1 PT in VISBO-Ressources
                                            Dim aktOrga As clsOrganisation = validOrganisations.getOrganisationValidAt(Date.Now)
                                            Dim hrole As New clsRolle(vgende - vganfang + 1)
                                            Dim otherRoledef As clsRollenDefinition
                                            Dim roledef As clsRollenDefinition
                                            If (aktOrga.allRoles.containsName(vgphase.verantwortlich)) Then
                                                otherRoledef = RoleDefinitions.getRoledef(vgphase.verantwortlich)
                                                roledef = aktOrga.allRoles.getRoledef(vgphase.verantwortlich)
                                            Else
                                                Dim defaultTopNode As String = RoleDefinitions.getDefaultTopNodeName()
                                                otherRoledef = RoleDefinitions.getRoledef(defaultTopNode)
                                                roledef = aktOrga.allRoles.getRoledef(defaultTopNode)
                                            End If

                                            hrole.uid = roledef.UID
                                            hrole.teamID = -1

                                            ReDim vgoldXwerte(0)
                                            vgoldXwerte(0) = itemVg.Value.StoryPoints

                                            With vgphase
                                                ReDim vgXwerte(vgende - vganfang + 1)
                                                .berechneBedarfe(.getStartDate, .getEndDate, vgoldXwerte, 1, vgXwerte)
                                            End With
                                            hrole.Xwerte = vgXwerte

                                            vgphase.addRole(hrole)

                                        End If

                                        ' PMO schreibt BaseLine, die aber nur die Epics enthalten soll
                                        If myCustomUserRole.customUserRole = ptCustomUserRoles.PortfolioManager Then

                                            tasksInserted.Add(itemVg.Value.Jira_ID, itemVg.Value)
                                            ' PMO schreibt BaseLine, die aber nur die Epics enthalten soll
                                            ephase.unionizeWith(vgphase)

                                        Else
                                            ' hphase in Hierarchie auf Level 1 eintragen und in Projekt einhängen
                                            Dim hrchynode As New clsHierarchyNode
                                            hrchynode.elemName = vgphase.name
                                            hrchynode.parentNodeKey = ephase.nameID
                                            hproj.AddPhase(vgphase, origName:=itemVg.Value.Jira_ID, parentID:=ephase.nameID)
                                            tasksInserted.Add(itemVg.Value.Jira_ID, itemVg.Value)
                                            hrchynode.indexOfElem = hproj.AllPhases.Count
                                        End If

                                    End If

                                    ' bestimme retrospektiv phaseStart und phaseEnd
                                    If vgphase.getStartDate < phaseStart Then
                                        phaseStart = vgphase.getStartDate
                                    End If
                                    If vgphase.getEndDate > phaseEnd Then
                                        phaseEnd = vgphase.getEndDate
                                    End If
                                    ' hier werden die Anzahl erledigter Issues je epic gezählt
                                    If itemVg.Value.erledigt > Date.MinValue Then
                                        erledigteVgCount = erledigteVgCount + 1
                                    End If

                                    Call logger(ptErrLevel.logInfo, "JIRA-Task " & itemVg.Value.Jira_ID & ":" & itemVg.Value.Zusammenfassung & " gelesen", "readProjectsJIRA", anzFehler)

                                Next     ' itemvg = Vorgang



                                ' Dauer der Phase anpassen an Tasks-Dates
                                duration = calcDauerIndays(phaseStart, phaseEnd)
                                If duration > 0 Then
                                    Dim offset As Integer = DateDiff(DateInterval.Day, hproj.startDate.Date, phaseStart.Date)
                                    ephase.offset = offset
                                    ephase.changeStartandDauer(offset, duration)
                                End If
                                ' Dauer ist nun korrigiert

                                ' wieviel vorgänge des Epic erledigt sind in Prozent
                                If epicVorg.Count > 0 Then
                                    ephase.percentDone = erledigteVgCount / epicVorg.Count
                                Else
                                    ephase.percentDone = 0
                                End If

                                ' Ressources der Phase (= Summe der Tasks zu einem Epic) auf die Dauer verteilen
                                Dim Xwerte As Double() = Nothing
                                Dim oldXwerte As Double()
                                Dim anfang As Integer = ephase.relStart
                                Dim ende As Integer = ephase.relEnde

                                If epicStoryPoints > 0.0 Then
                                    ' ein StoryPoint in JIRA entspricht  1 PT in VISBO-Ressources
                                    Dim aktOrga As clsOrganisation = validOrganisations.getOrganisationValidAt(Date.Now)
                                    Dim hrole As New clsRolle(ende - anfang + 1)
                                    Dim otherRoledef As clsRollenDefinition
                                    Dim roledef As clsRollenDefinition

                                    If (aktOrga.allRoles.containsName(ephase.verantwortlich)) Then
                                        otherRoledef = RoleDefinitions.getRoledef(ephase.verantwortlich)
                                        roledef = aktOrga.allRoles.getRoledef(ephase.verantwortlich)
                                    Else
                                        Dim defaultTopNode As String = RoleDefinitions.getDefaultTopNodeName()
                                        otherRoledef = RoleDefinitions.getRoledef(defaultTopNode)
                                        roledef = aktOrga.allRoles.getRoledef(defaultTopNode)
                                    End If

                                    hrole.uid = roledef.UID
                                    hrole.teamID = -1

                                    ReDim oldXwerte(0)
                                    oldXwerte(0) = item.Value.StoryPoints

                                    With ephase
                                        ReDim Xwerte(ende - anfang)
                                        .berechneBedarfe(.getStartDate, .getEndDate, oldXwerte, 1, Xwerte)
                                    End With
                                    hrole.Xwerte = Xwerte

                                    ephase.addRole(hrole)

                                End If

                                ie = ie + 1
                                Call logger(ptErrLevel.logInfo, "JIRA-Phase " & item.Value.Jira_ID & ":" & item.Value.Zusammenfassung & " gelesen", "readProjectsJIRA", anzFehler)

                            Next    ' item = epic

                        Catch ex2 As Exception
                            Call logger(ptErrLevel.logError, "JIRA-Phase " & ex2.Message, "readProjectsJIRA", anzFehler)
                            Call MsgBox("line 4789")

                        End Try

                        'restliche Tasks des Projektes in die RootPhase eintragen und wenn einem Sprint zugeordnet
                        For Each task As KeyValuePair(Of String, clsJIRA_Task) In taskListSorted

                            If Not tasksInserted.ContainsKey(task.Key) Then
                                tasksRemaining.Add(task.Key, task.Value)
                                If (task.Value.SprintName = "" Or IsNothing(task.Value.SprintName)) And task.Value.TaskStatus <> "Fertig" Then
                                    tasksBacklog.Add(task.Key, task.Value)
                                ElseIf task.Value.TaskStatus = "Fertig" Then
                                    tasksFertigOSprint.Add(task.Key, task.Value)
                                End If
                            End If
                        Next  ' Ende bestimmen der restl. Tasks

                        ' Backlog bearbeiten, sofern es gibt

                        If tasksBacklog.Count <= 0 Then
                            Call logger(ptErrLevel.logInfo, "Es gibt keine Backlog-Tasks", "readProjectsJIRA", anzFehler)
                        Else
                            ' Backlog-Tasks eintragen last SprintEnde - ProjektEnde
                            ' Bestimmung von Start und Ende der Tasks - evt. durch last SprintEnde und ProjektEnde definiert
                            Dim bphase As clsPhase
                            'Dim testDate As Date = sprintList.ElementAt(sprintList.Count - 1).Value.SprintEndDate
                            Dim backStart As Date = lastSprintEnd
                            Dim backEnd As Date = hproj.endeDate

                            bphase = New clsPhase(hproj)
                            Dim backphaseNameNew As String = "Backlog without epics"
                            bphase.nameID = calcHryElemKey(backphaseNameNew, False)
                            Dim durationBL As Integer = 0

                            durationBL = calcDauerIndays(backStart, backEnd)

                            If durationBL > 0 Then
                                If backStart < hproj.startDate Then
                                    backStart = hproj.startDate
                                End If
                                Dim offset As Integer = DateDiff(DateInterval.Day, hproj.startDate.Date, backStart.Date)
                                bphase.offset = offset
                                bphase.changeStartandDauer(offset, durationBL)

                                ' hphase in Hierarchie auf Level 1 eintragen und in Projekt einhängen
                                Dim hrchynode As New clsHierarchyNode
                                hrchynode.elemName = bphase.name
                                hrchynode.parentNodeKey = rootPhaseName
                                hproj.AddPhase(bphase, origName:=backphaseNameNew, parentID:=rootPhaseName)
                                hrchynode.indexOfElem = hproj.AllPhases.Count
                            End If

                            Dim backphase As clsPhase
                            Dim backlogPhaseID As String = bphase.nameID         ' ist der Parent der Backlog - Tasks
                            For Each backlogItem As KeyValuePair(Of String, clsJIRA_Task) In tasksBacklog

                                backphase = New clsPhase(hproj)
                                backphaseNameNew = backlogItem.Value.Jira_ID & " " & backlogItem.Value.Zusammenfassung
                                backphase.nameID = calcHryElemKey(backphaseNameNew, False)

                                ' falls Synonyme definiert sind, ersetzen durch Std-Name, sonst bleibt Name unverändert 
                                Dim origPhNameVG As String = backphaseNameNew
                                backphaseNameNew = phaseMappings.mapToStdName("", backphaseNameNew)

                                If durationBL > 0 Then
                                    Dim offset As Integer = DateDiff(DateInterval.Day, hproj.startDate.Date, backStart.Date)
                                    backphase.offset = offset
                                    backphase.changeStartandDauer(offset, durationBL)
                                    If backlogItem.Value.TaskStatus = "Fertig" Then
                                        backphase.percentDone = 1.0
                                    Else
                                        backphase.percentDone = 0.0
                                    End If
                                    backphase.verantwortlich = backlogItem.Value.zugewPerson

                                    ' Ressources eintragen für Backlog-Tasks
                                    Dim Xwerte As Double() = Nothing
                                    Dim oldXwerte As Double()
                                    Dim anfang As Integer = backphase.relStart
                                    Dim ende As Integer = backphase.relEnde
                                    If backlogItem.Value.StoryPoints > 0.0 Then
                                        ' ein StoryPoint in JIRA entspricht  1 PT in VISBO-Ressources
                                        Dim aktOrga As clsOrganisation = validOrganisations.getOrganisationValidAt(Date.Now)

                                        Dim hrole As New clsRolle(ende - anfang)
                                        Dim otherRoledef As clsRollenDefinition
                                        Dim roledef As clsRollenDefinition

                                        If (aktOrga.allRoles.containsName(backphase.verantwortlich)) Then
                                            otherRoledef = RoleDefinitions.getRoledef(backphase.verantwortlich)
                                            roledef = aktOrga.allRoles.getRoledef(backphase.verantwortlich)
                                        Else
                                            Dim defaultTopNode As String = RoleDefinitions.getDefaultTopNodeName()
                                            otherRoledef = RoleDefinitions.getRoledef(defaultTopNode)
                                            roledef = aktOrga.allRoles.getRoledef(defaultTopNode)
                                        End If

                                        hrole.uid = roledef.UID
                                        hrole.teamID = -1

                                        ReDim oldXwerte(0)
                                        oldXwerte(0) = backlogItem.Value.StoryPoints

                                        With backphase
                                            ReDim Xwerte(ende - anfang)
                                            .berechneBedarfe(.getStartDate, .getEndDate, oldXwerte, 1, Xwerte)
                                        End With
                                        hrole.Xwerte = Xwerte

                                        backphase.addRole(hrole)
                                    End If

                                    ' PMO schreibt BaseLine, die aber nur die Epics enthalten soll
                                    If myCustomUserRole.customUserRole = ptCustomUserRoles.PortfolioManager Then

                                        ' PMO schreibt BaseLine, die aber nur die Epics enthalten soll
                                        bphase.unionizeWith(backphase)

                                    Else
                                        ' hphase in Hierarchie auf Level 1 eintragen und in Projekt einhängen
                                        Dim hrchynode As New clsHierarchyNode
                                        hrchynode.elemName = backphase.name
                                        hrchynode.parentNodeKey = rootPhaseName
                                        hproj.AddPhase(backphase, origName:=backlogItem.Value.Jira_ID, parentID:=backlogPhaseID)
                                        'tasksInserted.Add(backlogItem.Value.Jira_ID, backlogItem.Value)
                                        hrchynode.indexOfElem = hproj.AllPhases.Count
                                    End If

                                End If
                                Call logger(ptErrLevel.logInfo, "JIRA-Task " & backlogItem.Value.Jira_ID & ":" & backlogItem.Value.Zusammenfassung & " gelesen", "readProjectsJIRA", anzFehler)

                            Next     ' backlogItem
                        End If

                        If result Then

                            Dim keyStr As String = ""
                            Try
                                keyStr = calcProjektKey(hproj)
                                ImportProjekte.Add(hproj, updateCurrentConstellation:=False)

                                Call logger(ptErrLevel.logInfo, "Einlesen JIRA-Projekte erfolgt " & tmpDatei, "readProjectsJIRA", anzFehler)
                                ' hier: merken der erfolgreich importierten Projects Dateien
                                If Not listOfArchivFiles.Contains(tmpDatei) Then
                                    listOfArchivFiles.Add(tmpDatei)
                                End If


                            Catch ex2 As Exception
                                Call MsgBox("Projekt " & keyStr & " kann nicht zweimal importiert werden ...")
                            End Try

                        Else
                            Call logger(ptErrLevel.logWarning, "Einlesen JIRA-Projekt nicht erfolgt " & tmpDatei, "readProjectsJIRA", anzFehler)
                        End If

                    Next        ' for each kvp in projListSortedName

                End If      ' reading of all tasks ok

            Next    ' for each tmpDatei

        Else
            Dim errMsg As String = "Es gibt keine Datei zur Projekt-Anlage" & vbLf _
                             & "Es wurden daher jetzt keine berücksichtigt"

            ' das sollte nicht dazu führen, dass nichts gemacht wird 
            'meldungen.Add(errMsg)
            'ur: 08.01.2020: endgültige meldung erst nachdem alle abgearbeitet wurden
            'Call MsgBox(errMsg)

            Call logger(ptErrLevel.logError, errMsg, "", anzFehler)
        End If

        If result Then
            readProjectsJIRA = listOfArchivFiles
        Else
            readProjectsJIRA = New List(Of String)
        End If

    End Function

    Public Function filternNach(ByVal Field As String, ByVal testValue As String, ByVal taskList As SortedList(Of Date, clsJIRA_Task)) As SortedList(Of Date, clsJIRA_Task)


        Dim resCollection As New SortedList(Of Date, clsJIRA_Task)

        For Each item As KeyValuePair(Of Date, clsJIRA_Task) In taskList
            Select Case Field
                Case "Vorgangstyp"
                    If item.Value.Vorgangstyp = testValue Then
                        resCollection.Add(item.Value.Erstellt, item.Value)
                    End If
                Case "ZugewiesenePerson"
                    If item.Value.zugewPerson = testValue Then
                        resCollection.Add(item.Value.Erstellt, item.Value)
                    End If
                Case "Autor"
                    If item.Value.Autor = testValue Then
                        resCollection.Add(item.Value.Erstellt, item.Value)
                    End If
                Case "Prio"
                    If item.Value.Prio = testValue Then
                        resCollection.Add(item.Value.Erstellt, item.Value)
                    End If
                Case "Task-Status"
                    If item.Value.TaskStatus = testValue Then
                        resCollection.Add(item.Value.Erstellt, item.Value)
                    End If
                Case "Verknüpfte Vorgänge"
                    If item.Value.verknüpfte_JiraID = testValue Then
                        resCollection.Add(item.Value.Erstellt, item.Value)
                    End If
                Case "Area"
                    If item.Value.Area = testValue Then
                        resCollection.Add(item.Value.Erstellt, item.Value)
                    End If
                Case "Label"
                Case "Übergeordnet"
                    If item.Value.parent_JiraID = testValue Then
                        resCollection.Add(item.Value.Erstellt, item.Value)
                    End If
                Case "SprintName"
                    If item.Value.SprintName = testValue Then
                        resCollection.Add(item.Value.Erstellt, item.Value)
                    End If
                Case Else

            End Select
        Next

        filternNach = resCollection
    End Function


    Public Function readJIRATasks(ByVal projectConfig As SortedList(Of String, clsConfigProjectsImport),
                                  ByVal JiraExcelFile As String, ByRef projtaskList As SortedList(Of String, SortedList(Of Date, clsJIRA_Task)),
                                  ByRef projListSorted As SortedList(Of String, SortedList(Of String, clsJIRA_Task)),
                                  ByRef projsprintList As SortedList(Of String, SortedList(Of String, clsJIRA_sprint)), ByRef meldungen As Collection) As Boolean
        Dim result As Boolean = False
        Dim outputline As String = ""
        Dim ok As Boolean = False
        Dim projectWB As Microsoft.Office.Interop.Excel.Workbook = Nothing
        Dim currentWS As Microsoft.Office.Interop.Excel.Worksheet = Nothing
        Dim regexpression As Regex
        Dim firstUrlspalte As Integer
        Dim firstUrlzeile As Integer
        Dim lastSpalte As Integer
        Dim lastZeile As Integer
        Dim anz_Proj_created As Integer = 0
        Dim anz_Proj_notCreated As Integer = 0
        Dim oneNextTask As New clsJIRA_Task
        Dim taskListSorted As New SortedList(Of String, clsJIRA_Task)
        Dim taskList As New SortedList(Of Date, clsJIRA_Task)
        Dim sprintList As New SortedList(Of String, clsJIRA_sprint)

        Dim noGo As Integer = 0   'Sobald diese Variable > 0 ist, wird das Projekt nicht importiert

        Try
            If My.Computer.FileSystem.FileExists(JiraExcelFile) Then

                Try
                    projectWB = appInstance.Workbooks.Open(JiraExcelFile)

                    Dim vstart As clsConfigProjectsImport = projectConfig("valueStart")
                    ' Auslesen erste Projekt-Spalte
                    firstUrlspalte = vstart.column.von
                    firstUrlzeile = vstart.row.von

                    If appInstance.Worksheets.Count > 0 Then

                        If Not IsNothing(vstart.sheet) Then
                            currentWS = CType(appInstance.Worksheets(vstart.sheet), Global.Microsoft.Office.Interop.Excel.Worksheet)
                            ok = (currentWS.Name = vstart.sheetDescript)
                        End If
                        If Not ok Then
                            If Not IsNothing(vstart.sheetDescript) Then
                                currentWS = CType(appInstance.Worksheets(vstart.sheetDescript), Global.Microsoft.Office.Interop.Excel.Worksheet)
                            Else
                                currentWS = Nothing
                            End If
                        End If

                        If IsNothing(currentWS) Then
                            outputline = "The Worksheet you want to import cannot be matched"
                            meldungen.Add(outputline)
                            Call logger(ptErrLevel.logError, outputline, "readJIRATasks", anzFehler)
                        Else
                            ' bestimme letzte Zeile und Spalte
                            lastSpalte = CType(currentWS.Cells(firstUrlzeile, 2000), Global.Microsoft.Office.Interop.Excel.Range).End(Excel.XlDirection.xlToLeft).Column
                            lastZeile = CType(currentWS.Cells(2000, firstUrlspalte), Global.Microsoft.Office.Interop.Excel.Range).End(Excel.XlDirection.xlUp).Row

                            Try


                                For i = firstUrlzeile To lastZeile

                                    oneNextTask = New clsJIRA_Task

                                    ' projectName:
                                    Try
                                        Dim projectNameConfig As clsConfigProjectsImport = projectConfig("Projekt")
                                        'richtige Tabelle öffnen
                                        If currentWS.Index <> projectNameConfig.sheet Then
                                            If Not IsNothing(projectNameConfig.sheet) Then
                                                currentWS = CType(appInstance.Worksheets(projectNameConfig.sheet), Global.Microsoft.Office.Interop.Excel.Worksheet)
                                            Else
                                                currentWS = CType(appInstance.Worksheets(projectNameConfig.sheetDescript), Global.Microsoft.Office.Interop.Excel.Worksheet)
                                            End If
                                        End If
                                        With currentWS
                                            Select Case projectNameConfig.Typ
                                                Case "Text"
                                                    oneNextTask.projectName = CStr(.Cells(i, projectNameConfig.column.von).value)
                                                Case "Integer"
                                                    oneNextTask.projectName = CInt(.Cells(i, projectNameConfig.column.von).value)
                                                Case "Decimal"
                                                    oneNextTask.projectName = CDbl(.Cells(i, projectNameConfig.column.von).value)
                                                Case "Date"
                                                    oneNextTask.projectName = CDate(.Cells(i, projectNameConfig.column.von).value)
                                                Case Else
                                                    oneNextTask.projectName = .Cells(i, projectNameConfig.column.von).value
                                            End Select

                                            If projectNameConfig.objType = "RegEx" Then
                                                regexpression = New Regex(projectNameConfig.content)
                                                Dim match As Match = regexpression.Match(oneNextTask.projectName)
                                                If match.Success Then
                                                    oneNextTask.projectName = match.Value
                                                Else
                                                    oneNextTask.projectName = Nothing
                                                End If
                                            End If
                                        End With
                                    Catch ex As Exception
                                        ' Fehler bei ProjectName
                                        outputline = "problems reading ProjektName: line: " & i.ToString
                                        meldungen.Add(outputline)
                                        Call logger(ptErrLevel.logError, outputline, "readJIRATasks", anzFehler)
                                    End Try

                                    ' Vorgangstyp:
                                    Try
                                        Dim vorgangstypConfig As clsConfigProjectsImport = projectConfig("Vorgangstyp")
                                        'richtige Tabelle öffnen
                                        If currentWS.Index <> vorgangstypConfig.sheet Then
                                            If Not IsNothing(vorgangstypConfig.sheet) Then
                                                currentWS = CType(appInstance.Worksheets(vorgangstypConfig.sheet), Global.Microsoft.Office.Interop.Excel.Worksheet)
                                            Else
                                                currentWS = CType(appInstance.Worksheets(vorgangstypConfig.sheetDescript), Global.Microsoft.Office.Interop.Excel.Worksheet)
                                            End If
                                        End If
                                        With currentWS
                                            Select Case vorgangstypConfig.Typ
                                                Case "Text"
                                                    oneNextTask.Vorgangstyp = CStr(.Cells(i, vorgangstypConfig.column.von).value)
                                                Case "Integer"
                                                    oneNextTask.Vorgangstyp = CInt(.Cells(i, vorgangstypConfig.column.von).value)
                                                Case "Decimal"
                                                    oneNextTask.Vorgangstyp = CDbl(.Cells(i, vorgangstypConfig.column.von).value)
                                                Case "Date"
                                                    oneNextTask.Vorgangstyp = CDate(.Cells(i, vorgangstypConfig.column.von).value)
                                                Case Else
                                                    oneNextTask.Vorgangstyp = .Cells(i, vorgangstypConfig.column.von).value
                                            End Select

                                            If vorgangstypConfig.objType = "RegEx" Then
                                                regexpression = New Regex(vorgangstypConfig.content)
                                                Dim match As Match = regexpression.Match(oneNextTask.Vorgangstyp)
                                                If match.Success Then
                                                    oneNextTask.Vorgangstyp = match.Value
                                                Else
                                                    oneNextTask.Vorgangstyp = Nothing
                                                End If
                                            End If
                                        End With
                                    Catch ex As Exception
                                        ' Fehler bei Vorgangstyp
                                        outputline = "problems reading Vorgangstyp: line: " & i.ToString
                                        meldungen.Add(outputline)
                                        Call logger(ptErrLevel.logError, outputline, "readJIRATasks", anzFehler)
                                    End Try

                                    'Jira-ID:
                                    Try
                                        Dim JiraIDConfig As clsConfigProjectsImport = projectConfig("Jira-ID")
                                        'richtige Tabelle öffnen
                                        If currentWS.Index <> JiraIDConfig.sheet Then
                                            If Not IsNothing(JiraIDConfig.sheet) Then
                                                currentWS = CType(appInstance.Worksheets(JiraIDConfig.sheet), Global.Microsoft.Office.Interop.Excel.Worksheet)
                                            Else
                                                currentWS = CType(appInstance.Worksheets(JiraIDConfig.sheetDescript), Global.Microsoft.Office.Interop.Excel.Worksheet)
                                            End If
                                        End If
                                        With currentWS
                                            Select Case JiraIDConfig.Typ
                                                Case "Text"
                                                    oneNextTask.Jira_ID = CStr(.Cells(i, JiraIDConfig.column.von).value)
                                                Case "Integer"
                                                    oneNextTask.Jira_ID = CInt(.Cells(i, JiraIDConfig.column.von).value)
                                                Case "Decimal"
                                                    oneNextTask.Jira_ID = CDbl(.Cells(i, JiraIDConfig.column.von).value)
                                                Case "Date"
                                                    oneNextTask.Jira_ID = CDate(.Cells(i, JiraIDConfig.column.von).value)
                                                Case Else
                                                    oneNextTask.Jira_ID = .Cells(i, JiraIDConfig.column.von).value
                                            End Select

                                            If JiraIDConfig.objType = "RegEx" Then
                                                regexpression = New Regex(JiraIDConfig.content)
                                                Dim match As Match = regexpression.Match(oneNextTask.Jira_ID)
                                                If match.Success Then
                                                    oneNextTask.Jira_ID = match.Value
                                                Else
                                                    oneNextTask.Jira_ID = Nothing
                                                End If
                                            End If
                                        End With

                                    Catch ex As Exception
                                        ' Fehler bei Jira-ID:
                                        outputline = "problems reading JIRA-ID: line: " & i.ToString
                                        meldungen.Add(outputline)
                                        Call logger(ptErrLevel.logError, outputline, "readJIRATasks", anzFehler)
                                    End Try

                                    'Zusammenfassung:
                                    Try
                                        Dim subjectConfig As clsConfigProjectsImport = projectConfig("Zusammenfassung")
                                        'richtige Tabelle öffnen
                                        If currentWS.Index <> subjectConfig.sheet Then
                                            If Not IsNothing(subjectConfig.sheet) Then
                                                currentWS = CType(appInstance.Worksheets(subjectConfig.sheet), Global.Microsoft.Office.Interop.Excel.Worksheet)
                                            Else
                                                currentWS = CType(appInstance.Worksheets(subjectConfig.sheetDescript), Global.Microsoft.Office.Interop.Excel.Worksheet)
                                            End If
                                        End If
                                        With currentWS
                                            Select Case subjectConfig.Typ
                                                Case "Text"
                                                    oneNextTask.Zusammenfassung = CStr(.Cells(i, subjectConfig.column.von).value)
                                                Case "Integer"
                                                    oneNextTask.Zusammenfassung = CInt(.Cells(i, subjectConfig.column.von).value)
                                                Case "Decimal"
                                                    oneNextTask.Zusammenfassung = CDbl(.Cells(i, subjectConfig.column.von).value)
                                                Case "Date"
                                                    oneNextTask.Zusammenfassung = CDate(.Cells(i, subjectConfig.column.von).value)
                                                Case Else
                                                    oneNextTask.Zusammenfassung = .Cells(i, subjectConfig.column.von).value
                                            End Select

                                            If subjectConfig.objType = "RegEx" Then
                                                regexpression = New Regex(subjectConfig.content)
                                                Dim match As Match = regexpression.Match(oneNextTask.Zusammenfassung)
                                                If match.Success Then
                                                    oneNextTask.Zusammenfassung = match.Value
                                                Else
                                                    oneNextTask.Zusammenfassung = Nothing
                                                End If
                                            End If
                                        End With

                                    Catch ex As Exception
                                        ' Fehler bei Zusammenfassung:
                                        outputline = "problems reading Zusammenfassung: line: " & i.ToString
                                        meldungen.Add(outputline)
                                        Call logger(ptErrLevel.logError, outputline, "readJIRATasks", anzFehler)
                                    End Try

                                    'zugewiesene Person:
                                    Try
                                        Dim personConfig As clsConfigProjectsImport = projectConfig("ZugewiesenePerson")
                                        'richtige Tabelle öffnen
                                        If currentWS.Index <> personConfig.sheet Then

                                            If Not IsNothing(personConfig.sheet) Then
                                                currentWS = CType(appInstance.Worksheets(personConfig.sheet), Global.Microsoft.Office.Interop.Excel.Worksheet)
                                            Else
                                                currentWS = CType(appInstance.Worksheets(personConfig.sheetDescript), Global.Microsoft.Office.Interop.Excel.Worksheet)
                                            End If
                                        End If
                                        With currentWS
                                            Select Case personConfig.Typ
                                                Case "Text"
                                                    oneNextTask.zugewPerson = CStr(.Cells(i, personConfig.column.von).value)
                                                Case "Integer"
                                                    oneNextTask.zugewPerson = CInt(.Cells(i, personConfig.column.von).value)
                                                Case "Decimal"
                                                    oneNextTask.zugewPerson = CDbl(.Cells(i, personConfig.column.von).value)
                                                Case "Date"
                                                    oneNextTask.zugewPerson = CDate(.Cells(i, personConfig.column.von).value)
                                                Case Else
                                                    oneNextTask.zugewPerson = .Cells(i, personConfig.column.von).value
                                            End Select

                                            If personConfig.objType = "RegEx" Then
                                                regexpression = New Regex(personConfig.content)
                                                Dim match As Match = regexpression.Match(oneNextTask.zugewPerson)
                                                If match.Success Then
                                                    oneNextTask.zugewPerson = match.Value
                                                Else
                                                    oneNextTask.zugewPerson = Nothing
                                                End If
                                            End If
                                        End With

                                    Catch ex As Exception
                                        ' Fehler bei zugewiesene Person:
                                        outputline = "problems reading zugewiesene Person: line: " & i.ToString
                                        meldungen.Add(outputline)
                                        Call logger(ptErrLevel.logError, outputline, "readJIRATasks", anzFehler)
                                    End Try

                                    'Autor:
                                    Try
                                        Dim autorConfig As clsConfigProjectsImport = projectConfig("Autor")
                                        'richtige Tabelle öffnen
                                        If currentWS.Index <> autorConfig.sheet Then
                                            If Not IsNothing(autorConfig.sheet) Then
                                                currentWS = CType(appInstance.Worksheets(autorConfig.sheet), Global.Microsoft.Office.Interop.Excel.Worksheet)
                                            Else
                                                currentWS = CType(appInstance.Worksheets(autorConfig.sheetDescript), Global.Microsoft.Office.Interop.Excel.Worksheet)
                                            End If
                                        End If
                                        With currentWS
                                            Select Case autorConfig.Typ
                                                Case "Text"
                                                    oneNextTask.Autor = CStr(.Cells(i, autorConfig.column.von).value)
                                                Case "Integer"
                                                    oneNextTask.Autor = CInt(.Cells(i, autorConfig.column.von).value)
                                                Case "Decimal"
                                                    oneNextTask.Autor = CDbl(.Cells(i, autorConfig.column.von).value)
                                                Case "Date"
                                                    oneNextTask.Autor = CDate(.Cells(i, autorConfig.column.von).value)
                                                Case Else
                                                    oneNextTask.Autor = .Cells(i, autorConfig.column.von).value
                                            End Select

                                            If autorConfig.objType = "RegEx" Then
                                                regexpression = New Regex(autorConfig.content)
                                                Dim match As Match = regexpression.Match(oneNextTask.Autor)
                                                If match.Success Then
                                                    oneNextTask.Autor = match.Value
                                                Else
                                                    oneNextTask.Autor = Nothing
                                                End If
                                            End If
                                        End With

                                    Catch ex As Exception
                                        ' Fehler bei zugewiesene Person:
                                        outputline = "problems reading Autor: line: " & i.ToString
                                        meldungen.Add(outputline)
                                        Call logger(ptErrLevel.logError, outputline, "readJIRATasks", anzFehler)
                                    End Try

                                    ' Priorität:
                                    Try
                                        Dim prioConfig As clsConfigProjectsImport = projectConfig("Prio")
                                        'richtige Tabelle öffnen
                                        If currentWS.Index <> prioConfig.sheet Then
                                            If Not IsNothing(prioConfig.sheet) Then
                                                currentWS = CType(appInstance.Worksheets(prioConfig.sheet), Global.Microsoft.Office.Interop.Excel.Worksheet)
                                            Else
                                                currentWS = CType(appInstance.Worksheets(prioConfig.sheetDescript), Global.Microsoft.Office.Interop.Excel.Worksheet)
                                            End If
                                        End If
                                        With currentWS
                                            Select Case prioConfig.Typ
                                                Case "Text"
                                                    oneNextTask.Prio = CStr(.Cells(i, prioConfig.column.von).value)
                                                Case "Integer"
                                                    oneNextTask.Prio = CInt(.Cells(i, prioConfig.column.von).value)
                                                Case "Decimal"
                                                    oneNextTask.Prio = CDbl(.Cells(i, prioConfig.column.von).value)
                                                Case "Date"
                                                    oneNextTask.Prio = CDate(.Cells(i, prioConfig.column.von).value)
                                                Case Else
                                                    oneNextTask.Prio = .Cells(i, prioConfig.column.von).value
                                            End Select

                                            If prioConfig.objType = "RegEx" Then
                                                regexpression = New Regex(prioConfig.content)
                                                Dim match As Match = regexpression.Match(oneNextTask.Prio)
                                                If match.Success Then
                                                    oneNextTask.Prio = match.Value
                                                Else
                                                    oneNextTask.Prio = Nothing
                                                End If
                                            End If
                                        End With

                                    Catch ex As Exception
                                        ' Fehler bei Priorität:
                                        outputline = "problems reading Priorität: line: " & i.ToString
                                        meldungen.Add(outputline)
                                        Call logger(ptErrLevel.logError, outputline, "readJIRATasks", anzFehler)
                                    End Try

                                    ' Task-Status:
                                    Try
                                        Dim statusConfig As clsConfigProjectsImport = projectConfig("Task-Status")
                                        'richtige Tabelle öffnen
                                        If currentWS.Index <> statusConfig.sheet Then
                                            If Not IsNothing(statusConfig.sheet) Then
                                                currentWS = CType(appInstance.Worksheets(statusConfig.sheet), Global.Microsoft.Office.Interop.Excel.Worksheet)
                                            Else
                                                currentWS = CType(appInstance.Worksheets(statusConfig.sheetDescript), Global.Microsoft.Office.Interop.Excel.Worksheet)
                                            End If
                                        End If
                                        With currentWS
                                            Select Case statusConfig.Typ
                                                Case "Text"
                                                    oneNextTask.TaskStatus = CStr(.Cells(i, statusConfig.column.von).value)
                                                Case "Integer"
                                                    oneNextTask.TaskStatus = CInt(.Cells(i, statusConfig.column.von).value)
                                                Case "Decimal"
                                                    oneNextTask.TaskStatus = CDbl(.Cells(i, statusConfig.column.von).value)
                                                Case "Date"
                                                    oneNextTask.TaskStatus = CDate(.Cells(i, statusConfig.column.von).value)
                                                Case Else
                                                    oneNextTask.TaskStatus = .Cells(i, statusConfig.column.von).value
                                            End Select

                                            If statusConfig.objType = "RegEx" Then
                                                regexpression = New Regex(statusConfig.content)
                                                Dim match As Match = regexpression.Match(oneNextTask.TaskStatus)
                                                If match.Success Then
                                                    oneNextTask.TaskStatus = match.Value
                                                Else
                                                    oneNextTask.TaskStatus = Nothing
                                                End If
                                            End If
                                        End With

                                    Catch ex As Exception
                                        ' Fehler bei Task-Status:
                                        outputline = "problems reading Status: line: " & i.ToString
                                        meldungen.Add(outputline)
                                        Call logger(ptErrLevel.logError, outputline, "readJIRATasks", anzFehler)
                                    End Try

                                    ' Lösung:
                                    Try
                                        Dim lösungConfig As clsConfigProjectsImport = projectConfig("Lösung(*)")
                                        'richtige Tabelle öffnen
                                        If currentWS.Index <> lösungConfig.sheet Then
                                            If Not IsNothing(lösungConfig.sheet) Then
                                                currentWS = CType(appInstance.Worksheets(lösungConfig.sheet), Global.Microsoft.Office.Interop.Excel.Worksheet)
                                            Else
                                                currentWS = CType(appInstance.Worksheets(lösungConfig.sheetDescript), Global.Microsoft.Office.Interop.Excel.Worksheet)
                                            End If
                                        End If
                                        With currentWS
                                            Select Case lösungConfig.Typ
                                                Case "Text"
                                                    oneNextTask.loesung = CStr(.Cells(i, lösungConfig.column.von).value)
                                                Case "Integer"
                                                    oneNextTask.loesung = CInt(.Cells(i, lösungConfig.column.von).value)
                                                Case "Decimal"
                                                    oneNextTask.loesung = CDbl(.Cells(i, lösungConfig.column.von).value)
                                                Case "Date"
                                                    oneNextTask.loesung = CDate(.Cells(i, lösungConfig.column.von).value)
                                                Case Else
                                                    oneNextTask.loesung = .Cells(i, lösungConfig.column.von).value
                                            End Select

                                            If lösungConfig.objType = "RegEx" Then
                                                regexpression = New Regex(lösungConfig.content)
                                                Dim match As Match = regexpression.Match(oneNextTask.loesung)
                                                If match.Success Then
                                                    oneNextTask.loesung = match.Value
                                                Else
                                                    oneNextTask.loesung = Nothing
                                                End If
                                            End If
                                        End With

                                    Catch ex As Exception
                                        ' Fehler bei Lösung(*)
                                        outputline = "problems reading Lösung: line: " & i.ToString
                                        meldungen.Add(outputline)
                                        Call logger(ptErrLevel.logError, outputline, "readJIRATasks", anzFehler)
                                    End Try

                                    ' Erstellungsdatum:
                                    Try
                                        Dim createdAtConfig As clsConfigProjectsImport = projectConfig("Erstellungsdatum")
                                        'richtige Tabelle öffnen
                                        If currentWS.Index <> createdAtConfig.sheet Then
                                            If Not IsNothing(createdAtConfig.sheet) Then
                                                currentWS = CType(appInstance.Worksheets(createdAtConfig.sheet), Global.Microsoft.Office.Interop.Excel.Worksheet)
                                            Else
                                                currentWS = CType(appInstance.Worksheets(createdAtConfig.sheetDescript), Global.Microsoft.Office.Interop.Excel.Worksheet)
                                            End If
                                        End If
                                        With currentWS
                                            Select Case createdAtConfig.Typ
                                                Case "Text"
                                                    oneNextTask.Erstellt = CStr(.Cells(i, createdAtConfig.column.von).value)
                                                'Case "Integer"
                                                '    oneNextTask.Erstellt = CInt(.Cells(i, createedAtConfig.column.von).value)
                                                'Case "Decimal"
                                                '    oneNextTask.Erstellt = CDbl(.Cells(i, createedAtConfig.column.von).value)
                                                Case "Date"
                                                    oneNextTask.Erstellt = CDate(.Cells(i, createdAtConfig.column.von).value)
                                                Case Else
                                                    oneNextTask.Erstellt = .Cells(i, createdAtConfig.column.von).value
                                            End Select

                                            If createdAtConfig.objType = "RegEx" Then
                                                regexpression = New Regex(createdAtConfig.content)
                                                Dim match As Match = regexpression.Match(oneNextTask.Erstellt)
                                                If match.Success Then
                                                    oneNextTask.Erstellt = match.Value
                                                Else
                                                    oneNextTask.Erstellt = Nothing
                                                End If
                                            End If
                                        End With

                                    Catch ex As Exception
                                        ' Fehler bei Erstellungsdatum
                                        outputline = "problems reading Erstellungsdatum: line: " & i.ToString
                                        meldungen.Add(outputline)
                                        Call logger(ptErrLevel.logError, outputline, "readJIRATasks", anzFehler)
                                    End Try

                                    ' Aktualisiert:
                                    Try
                                        Dim updatedAtConfig As clsConfigProjectsImport = projectConfig("Aktualisiert")
                                        'richtige Tabelle öffnen
                                        If currentWS.Index <> updatedAtConfig.sheet Then
                                            If Not IsNothing(updatedAtConfig.sheet) Then
                                                currentWS = CType(appInstance.Worksheets(updatedAtConfig.sheet), Global.Microsoft.Office.Interop.Excel.Worksheet)
                                            Else
                                                currentWS = CType(appInstance.Worksheets(updatedAtConfig.sheetDescript), Global.Microsoft.Office.Interop.Excel.Worksheet)
                                            End If
                                        End If
                                        With currentWS
                                            Select Case updatedAtConfig.Typ
                                                Case "Text"
                                                    oneNextTask.aktualisiert = CStr(.Cells(i, updatedAtConfig.column.von).value)
                                                'Case "Integer"
                                                '    oneNextTask.aktualisiert = CInt(.Cells(i, createedAtConfig.column.von).value)
                                                'Case "Decimal"
                                                '    oneNextTask.aktualisiert = CDbl(.Cells(i, createedAtConfig.column.von).value)
                                                Case "Date"
                                                    oneNextTask.aktualisiert = CDate(.Cells(i, updatedAtConfig.column.von).value)
                                                Case Else
                                                    oneNextTask.aktualisiert = .Cells(i, updatedAtConfig.column.von).value
                                            End Select

                                            If updatedAtConfig.objType = "RegEx" Then
                                                regexpression = New Regex(updatedAtConfig.content)
                                                Dim match As Match = regexpression.Match(oneNextTask.aktualisiert)
                                                If match.Success Then
                                                    oneNextTask.aktualisiert = match.Value
                                                Else
                                                    oneNextTask.aktualisiert = Nothing
                                                End If
                                            End If
                                        End With

                                    Catch ex As Exception
                                        ' Fehler bei Aktualisiert
                                        outputline = "problems reading Aktualisierungs-Datum: line: " & i.ToString
                                        meldungen.Add(outputline)
                                        Call logger(ptErrLevel.logError, outputline, "readJIRATasks", anzFehler)
                                    End Try

                                    ' Fälligkeitsdatum:
                                    Try
                                        Dim toBeReadyConfig As clsConfigProjectsImport = projectConfig("Fälligkeitsdatum")
                                        'richtige Tabelle öffnen
                                        If currentWS.Index <> toBeReadyConfig.sheet Then
                                            If Not IsNothing(toBeReadyConfig.sheet) Then
                                                currentWS = CType(appInstance.Worksheets(toBeReadyConfig.sheet), Global.Microsoft.Office.Interop.Excel.Worksheet)
                                            Else
                                                currentWS = CType(appInstance.Worksheets(toBeReadyConfig.sheetDescript), Global.Microsoft.Office.Interop.Excel.Worksheet)
                                            End If
                                        End If
                                        With currentWS
                                            Select Case toBeReadyConfig.Typ
                                                Case "Text"
                                                    oneNextTask.fällig = CStr(.Cells(i, toBeReadyConfig.column.von).value)
                                                'Case "Integer"
                                                '    oneNextTask.fällig = CInt(.Cells(i, toBeReadyConfig.column.von).value)
                                                'Case "Decimal"
                                                '    oneNextTask.fällig = CDbl(.Cells(i, toBeReadyConfig.column.von).value)
                                                Case "Date"
                                                    oneNextTask.fällig = CDate(.Cells(i, toBeReadyConfig.column.von).value)
                                                Case Else
                                                    oneNextTask.fällig = .Cells(i, toBeReadyConfig.column.von).value
                                            End Select

                                            If toBeReadyConfig.objType = "RegEx" Then
                                                regexpression = New Regex(toBeReadyConfig.content)
                                                Dim match As Match = regexpression.Match(oneNextTask.fällig)
                                                If match.Success Then
                                                    oneNextTask.fällig = match.Value
                                                Else
                                                    oneNextTask.fällig = Nothing
                                                End If
                                            End If
                                        End With

                                    Catch ex As Exception
                                        ' Fehler bei Fälligkeitsdatum
                                        outputline = "problems reading Fälligkeit: line: " & i.ToString
                                        meldungen.Add(outputline)
                                        Call logger(ptErrLevel.logError, outputline, "readJIRATasks", anzFehler)
                                    End Try

                                    ' StartDatum:
                                    Try
                                        Dim startDateConfig As clsConfigProjectsImport = projectConfig("StartDate")
                                        'richtige Tabelle öffnen
                                        If currentWS.Index <> startDateConfig.sheet Then
                                            If Not IsNothing(startDateConfig.sheet) Then
                                                currentWS = CType(appInstance.Worksheets(startDateConfig.sheet), Global.Microsoft.Office.Interop.Excel.Worksheet)
                                            Else
                                                currentWS = CType(appInstance.Worksheets(startDateConfig.sheetDescript), Global.Microsoft.Office.Interop.Excel.Worksheet)
                                            End If
                                        End If
                                        With currentWS
                                            Select Case startDateConfig.Typ
                                                Case "Text"
                                                    oneNextTask.StartDate = CStr(.Cells(i, startDateConfig.column.von).value)
                                                'Case "Integer"
                                                '    oneNextTask.StartDate = CInt(.Cells(i, toBeReadyConfig.column.von).value)
                                                'Case "Decimal"
                                                '    oneNextTask.StartDate = CDbl(.Cells(i, toBeReadyConfig.column.von).value)
                                                Case "Date"
                                                    oneNextTask.StartDate = CDate(.Cells(i, startDateConfig.column.von).value)
                                                Case Else
                                                    oneNextTask.StartDate = .Cells(i, startDateConfig.column.von).value
                                            End Select

                                            If startDateConfig.objType = "RegEx" Then
                                                regexpression = New Regex(startDateConfig.content)
                                                Dim match As Match = regexpression.Match(oneNextTask.StartDate)
                                                If match.Success Then
                                                    oneNextTask.StartDate = match.Value
                                                Else
                                                    oneNextTask.StartDate = Nothing
                                                End If
                                            End If
                                        End With

                                    Catch ex As Exception
                                        ' Fehler bei Startdate
                                        outputline = "problems reading Start Date: line: " & i.ToString
                                        meldungen.Add(outputline)
                                        Call logger(ptErrLevel.logError, outputline, "readJIRATasks", anzFehler)
                                    End Try

                                    ' Verknüpfte Vorgänge(Jira-ID):
                                    Try
                                        Dim connectedConfig As clsConfigProjectsImport = projectConfig("Verknüpfte Vorgänge(Jira-ID)")
                                        'richtige Tabelle öffnen
                                        If currentWS.Index <> connectedConfig.sheet Then
                                            If Not IsNothing(connectedConfig.sheet) Then
                                                currentWS = CType(appInstance.Worksheets(connectedConfig.sheet), Global.Microsoft.Office.Interop.Excel.Worksheet)
                                            Else
                                                currentWS = CType(appInstance.Worksheets(connectedConfig.sheetDescript), Global.Microsoft.Office.Interop.Excel.Worksheet)
                                            End If
                                        End If
                                        With currentWS
                                            Select Case connectedConfig.Typ
                                                Case "Text"
                                                    oneNextTask.verknüpfte_JiraID = CStr(.Cells(i, connectedConfig.column.von).value)
                                                Case "Integer"
                                                    oneNextTask.verknüpfte_JiraID = CInt(.Cells(i, connectedConfig.column.von).value)
                                                Case "Decimal"
                                                    oneNextTask.verknüpfte_JiraID = CDbl(.Cells(i, connectedConfig.column.von).value)
                                                Case "Date"
                                                    oneNextTask.verknüpfte_JiraID = CDate(.Cells(i, connectedConfig.column.von).value)
                                                Case Else
                                                    oneNextTask.verknüpfte_JiraID = .Cells(i, connectedConfig.column.von).value
                                            End Select

                                            If connectedConfig.objType = "RegEx" Then
                                                regexpression = New Regex(connectedConfig.content)
                                                Dim match As Match = regexpression.Match(oneNextTask.verknüpfte_JiraID)
                                                If match.Success Then
                                                    oneNextTask.verknüpfte_JiraID = match.Value
                                                Else
                                                    oneNextTask.verknüpfte_JiraID = Nothing
                                                End If
                                            End If
                                        End With

                                    Catch ex As Exception
                                        ' Fehler bei verknüpfte Vorgänge
                                        outputline = "problems reading verknüpfte Vorgänge: line: " & i.ToString
                                        meldungen.Add(outputline)
                                        Call logger(ptErrLevel.logError, outputline, "readJIRATasks", anzFehler)
                                    End Try

                                    ' Area:
                                    Try
                                        Dim areaConfig As clsConfigProjectsImport = projectConfig("Area")
                                        'richtige Tabelle öffnen
                                        If currentWS.Index <> areaConfig.sheet Then
                                            If Not IsNothing(areaConfig.sheet) Then
                                                currentWS = CType(appInstance.Worksheets(areaConfig.sheet), Global.Microsoft.Office.Interop.Excel.Worksheet)
                                            Else
                                                currentWS = CType(appInstance.Worksheets(areaConfig.sheetDescript), Global.Microsoft.Office.Interop.Excel.Worksheet)
                                            End If
                                        End If
                                        With currentWS
                                            Select Case areaConfig.Typ
                                                Case "Text"
                                                    oneNextTask.verknüpfte_JiraID = CStr(.Cells(i, areaConfig.column.von).value)
                                                Case "Integer"
                                                    oneNextTask.verknüpfte_JiraID = CInt(.Cells(i, areaConfig.column.von).value)
                                                Case "Decimal"
                                                    oneNextTask.verknüpfte_JiraID = CDbl(.Cells(i, areaConfig.column.von).value)
                                                Case "Date"
                                                    oneNextTask.verknüpfte_JiraID = CDate(.Cells(i, areaConfig.column.von).value)
                                                Case Else
                                                    oneNextTask.verknüpfte_JiraID = .Cells(i, areaConfig.column.von).value
                                            End Select

                                            If areaConfig.objType = "RegEx" Then
                                                regexpression = New Regex(areaConfig.content)
                                                Dim match As Match = regexpression.Match(oneNextTask.verknüpfte_JiraID)
                                                If match.Success Then
                                                    oneNextTask.verknüpfte_JiraID = match.Value
                                                Else
                                                    oneNextTask.verknüpfte_JiraID = Nothing
                                                End If
                                            End If
                                        End With

                                    Catch ex As Exception
                                        ' Fehler bei Area
                                        outputline = "problems reading Area: line: " & i.ToString
                                        meldungen.Add(outputline)
                                        Call logger(ptErrLevel.logError, outputline, "readJIRATasks", anzFehler)
                                    End Try

                                    ' Parent(Jira-ID):
                                    Try
                                        Dim parentConfig As clsConfigProjectsImport = projectConfig("Parent(Jira-ID)")
                                        'richtige Tabelle öffnen
                                        If currentWS.Index <> parentConfig.sheet Then
                                            If Not IsNothing(parentConfig.sheet) Then
                                                currentWS = CType(appInstance.Worksheets(parentConfig.sheet), Global.Microsoft.Office.Interop.Excel.Worksheet)
                                            Else
                                                currentWS = CType(appInstance.Worksheets(parentConfig.sheetDescript), Global.Microsoft.Office.Interop.Excel.Worksheet)
                                            End If
                                        End If
                                        With currentWS
                                            Select Case parentConfig.Typ
                                                Case "Text"
                                                    oneNextTask.parent_JiraID = CStr(.Cells(i, parentConfig.column.von).value)
                                                Case "Integer"
                                                    oneNextTask.parent_JiraID = CInt(.Cells(i, parentConfig.column.von).value)
                                                Case "Decimal"
                                                    oneNextTask.parent_JiraID = CDbl(.Cells(i, parentConfig.column.von).value)
                                                Case "Date"
                                                    oneNextTask.parent_JiraID = CDate(.Cells(i, parentConfig.column.von).value)
                                                Case Else
                                                    oneNextTask.parent_JiraID = .Cells(i, parentConfig.column.von).value
                                            End Select

                                            If parentConfig.objType = "RegEx" Then
                                                regexpression = New Regex(parentConfig.content)
                                                Dim match As Match = regexpression.Match(oneNextTask.parent_JiraID)
                                                If match.Success Then
                                                    oneNextTask.parent_JiraID = match.Value
                                                Else
                                                    oneNextTask.parent_JiraID = Nothing
                                                End If
                                            End If
                                        End With

                                    Catch ex As Exception
                                        ' Fehler bei übergeordnet
                                        outputline = "problems reading Übergeordnet: line: " & i.ToString
                                        meldungen.Add(outputline)
                                        Call logger(ptErrLevel.logError, outputline, "readJIRATasks", anzFehler)
                                    End Try

                                    ' Fortschritt:
                                    Try
                                        Dim fortschrittConfig As clsConfigProjectsImport = projectConfig("Fortschritt")
                                        'richtige Tabelle öffnen
                                        If currentWS.Index <> fortschrittConfig.sheet Then
                                            If Not IsNothing(fortschrittConfig.sheet) Then
                                                currentWS = CType(appInstance.Worksheets(fortschrittConfig.sheet), Global.Microsoft.Office.Interop.Excel.Worksheet)
                                            Else
                                                currentWS = CType(appInstance.Worksheets(fortschrittConfig.sheetDescript), Global.Microsoft.Office.Interop.Excel.Worksheet)
                                            End If
                                        End If
                                        With currentWS
                                            Select Case fortschrittConfig.Typ
                                                Case "Text"
                                                    oneNextTask.Fortschritt = CStr(.Cells(i, fortschrittConfig.column.von).value)
                                                Case "Integer"
                                                    oneNextTask.Fortschritt = CInt(.Cells(i, fortschrittConfig.column.von).value)
                                                Case "Decimal"
                                                    oneNextTask.Fortschritt = CDbl(.Cells(i, fortschrittConfig.column.von).value)
                                                    'Case "Date"
                                                    '    oneNextTask.Fortschritt = CDate(.Cells(i, fortschrittConfig.column.von).value)
                                                Case Else
                                                    oneNextTask.Fortschritt = .Cells(i, fortschrittConfig.column.von).value
                                            End Select

                                            If fortschrittConfig.objType = "RegEx" Then
                                                regexpression = New Regex(fortschrittConfig.content)
                                                Dim match As Match = regexpression.Match(oneNextTask.Fortschritt)
                                                If match.Success Then
                                                    oneNextTask.Fortschritt = match.Value
                                                Else
                                                    oneNextTask.Fortschritt = Nothing
                                                End If
                                            End If
                                        End With

                                    Catch ex As Exception
                                        ' Fehler bei Fortschritt
                                        outputline = "problems reading Fortschritt: line: " & i.ToString
                                        meldungen.Add(outputline)
                                        Call logger(ptErrLevel.logError, outputline, "readJIRATasks", anzFehler)
                                    End Try

                                    ' Story Point-Schätzung:
                                    Try
                                        Dim aufwandConfig As clsConfigProjectsImport = projectConfig("Story Point-Schätzung")
                                        'richtige Tabelle öffnen
                                        If currentWS.Index <> aufwandConfig.sheet Then
                                            If Not IsNothing(aufwandConfig.sheet) Then
                                                currentWS = CType(appInstance.Worksheets(aufwandConfig.sheet), Global.Microsoft.Office.Interop.Excel.Worksheet)
                                            Else
                                                currentWS = CType(appInstance.Worksheets(aufwandConfig.sheetDescript), Global.Microsoft.Office.Interop.Excel.Worksheet)
                                            End If
                                        End If
                                        With currentWS
                                            Select Case aufwandConfig.Typ
                                                Case "Text"
                                                    oneNextTask.StoryPoints = CStr(.Cells(i, aufwandConfig.column.von).value)
                                                Case "Integer"
                                                    oneNextTask.StoryPoints = CInt(.Cells(i, aufwandConfig.column.von).value)
                                                Case "Decimal"
                                                    oneNextTask.StoryPoints = CDbl(.Cells(i, aufwandConfig.column.von).value)
                                                    'Case "Date"
                                                    '    oneNextTask.StoryPoints = CDate(.Cells(i, aufwandConfig.column.von).value)
                                                Case Else
                                                    oneNextTask.StoryPoints = .Cells(i, aufwandConfig.column.von).value
                                            End Select

                                            If aufwandConfig.objType = "RegEx" Then
                                                regexpression = New Regex(aufwandConfig.content)
                                                Dim match As Match = regexpression.Match(oneNextTask.StoryPoints)
                                                If match.Success Then
                                                    oneNextTask.StoryPoints = match.Value
                                                Else
                                                    oneNextTask.StoryPoints = Nothing
                                                End If
                                            End If
                                        End With

                                    Catch ex As Exception
                                        ' Fehler bei  Story Point-Schätzung
                                        outputline = "problems reading Story Point-Schätzung: line: " & i.ToString
                                        meldungen.Add(outputline)
                                        Call logger(ptErrLevel.logError, outputline, "readJIRATasks", anzFehler)
                                    End Try

                                    ' Erledigt-Datum:
                                    Try
                                        Dim erledigtConfig As clsConfigProjectsImport = projectConfig("Erledigt-Datum")
                                        'richtige Tabelle öffnen
                                        If currentWS.Index <> erledigtConfig.sheet Then
                                            If Not IsNothing(erledigtConfig.sheet) Then
                                                currentWS = CType(appInstance.Worksheets(erledigtConfig.sheet), Global.Microsoft.Office.Interop.Excel.Worksheet)
                                            Else
                                                currentWS = CType(appInstance.Worksheets(erledigtConfig.sheetDescript), Global.Microsoft.Office.Interop.Excel.Worksheet)
                                            End If
                                        End If
                                        With currentWS
                                            Select Case erledigtConfig.Typ
                                                Case "Text"
                                                    oneNextTask.erledigt = CStr(.Cells(i, erledigtConfig.column.von).value)
                                                'Case "Integer"
                                                '    oneNextTask.erledigt = CInt(.Cells(i, erledigtConfig.column.von).value)
                                                'Case "Decimal"
                                                '    oneNextTask.erledigt = CDbl(.Cells(i, erledigtConfig.column.von).value)
                                                Case "Date"
                                                    oneNextTask.erledigt = CDate(.Cells(i, erledigtConfig.column.von).value)
                                                Case Else
                                                    oneNextTask.erledigt = .Cells(i, erledigtConfig.column.von).value
                                            End Select

                                            If erledigtConfig.objType = "RegEx" Then
                                                regexpression = New Regex(erledigtConfig.content)
                                                Dim match As Match = regexpression.Match(oneNextTask.erledigt)
                                                If match.Success Then
                                                    oneNextTask.erledigt = match.Value
                                                Else
                                                    oneNextTask.erledigt = Nothing
                                                End If
                                            End If
                                        End With

                                    Catch ex As Exception
                                        ' Fehler bei  Erledigt-Datum
                                        outputline = "problems reading Erledigt-Datum: line: " & i.ToString
                                        meldungen.Add(outputline)
                                        Call logger(ptErrLevel.logError, outputline, "readJIRATasks", anzFehler)
                                    End Try

                                    ' SprintName:
                                    Try
                                        Dim sprintNameConfig As clsConfigProjectsImport = projectConfig("SprintName")
                                        'richtige Tabelle öffnen
                                        If currentWS.Index <> sprintNameConfig.sheet Then
                                            If Not IsNothing(sprintNameConfig.sheet) Then
                                                currentWS = CType(appInstance.Worksheets(sprintNameConfig.sheet), Global.Microsoft.Office.Interop.Excel.Worksheet)
                                            Else
                                                currentWS = CType(appInstance.Worksheets(sprintNameConfig.sheetDescript), Global.Microsoft.Office.Interop.Excel.Worksheet)
                                            End If
                                        End If
                                        With currentWS
                                            Select Case sprintNameConfig.Typ
                                                Case "Text"
                                                    oneNextTask.SprintName = CStr(.Cells(i, sprintNameConfig.column.von).value)
                                                Case "Integer"
                                                    oneNextTask.SprintName = CInt(.Cells(i, sprintNameConfig.column.von).value)
                                                Case "Decimal"
                                                    oneNextTask.SprintName = CDbl(.Cells(i, sprintNameConfig.column.von).value)
                                                Case "Date"
                                                    oneNextTask.SprintName = CDate(.Cells(i, sprintNameConfig.column.von).value)
                                                Case Else
                                                    oneNextTask.SprintName = .Cells(i, sprintNameConfig.column.von).value
                                            End Select

                                            If sprintNameConfig.objType = "RegEx" Then
                                                regexpression = New Regex(sprintNameConfig.content)
                                                Dim match As Match = regexpression.Match(oneNextTask.SprintName)
                                                If match.Success Then
                                                    oneNextTask.SprintName = match.Value
                                                Else
                                                    oneNextTask.SprintName = Nothing
                                                End If
                                            End If
                                        End With

                                    Catch ex As Exception
                                        ' Fehler bei  SprintName
                                        outputline = "problems reading SprintName: line: " & i.ToString
                                        meldungen.Add(outputline)
                                        Call logger(ptErrLevel.logError, outputline, "readJIRATasks", anzFehler)
                                    End Try

                                    ' SprintStartDate:
                                    Try
                                        Dim sprintStartConfig As clsConfigProjectsImport = projectConfig("SprintStartDate")
                                        'richtige Tabelle öffnen
                                        If currentWS.Index <> sprintStartConfig.sheet Then
                                            If Not IsNothing(sprintStartConfig.sheet) Then
                                                currentWS = CType(appInstance.Worksheets(sprintStartConfig.sheet), Global.Microsoft.Office.Interop.Excel.Worksheet)
                                            Else
                                                currentWS = CType(appInstance.Worksheets(sprintStartConfig.sheetDescript), Global.Microsoft.Office.Interop.Excel.Worksheet)
                                            End If
                                        End If
                                        With currentWS
                                            Select Case sprintStartConfig.Typ
                                                Case "Text"
                                                    oneNextTask.SprintStartDate = CStr(.Cells(i, sprintStartConfig.column.von).value)
                                                'Case "Integer"
                                                '    oneNextTask.SprintStartDate = CInt(.Cells(i, sprintStartConfig.column.von).value)
                                                'Case "Decimal"
                                                '    oneNextTask.SprintStartDate = CDbl(.Cells(i, sprintStartConfig.column.von).value)
                                                Case "Date"
                                                    oneNextTask.SprintStartDate = CDate(.Cells(i, sprintStartConfig.column.von).value)
                                                Case Else
                                                    oneNextTask.SprintStartDate = .Cells(i, sprintStartConfig.column.von).value
                                            End Select

                                            If sprintStartConfig.objType = "RegEx" Then
                                                regexpression = New Regex(sprintStartConfig.content)
                                                Dim match As Match = regexpression.Match(oneNextTask.SprintStartDate)
                                                If match.Success Then
                                                    oneNextTask.SprintStartDate = match.Value
                                                Else
                                                    oneNextTask.SprintStartDate = Nothing
                                                End If
                                            End If
                                        End With

                                    Catch ex As Exception
                                        ' Fehler bei  SprintStartDate
                                        oneNextTask.SprintStartDate = Nothing
                                        outputline = "problems reading SprintStartDate: line: " & i.ToString
                                        'meldungen.Add(outputline)
                                        Call logger(ptErrLevel.logError, outputline, "readJIRATasks", anzFehler)
                                    End Try


                                    ' SprintEndDate:
                                    Try
                                        Dim sprintEndConfig As clsConfigProjectsImport = projectConfig("SprintEndDate")
                                        'richtige Tabelle öffnen

                                        If currentWS.Index <> sprintEndConfig.sheet Then
                                            If Not IsNothing(sprintEndConfig.sheet) Then
                                                currentWS = CType(appInstance.Worksheets(sprintEndConfig.sheet), Global.Microsoft.Office.Interop.Excel.Worksheet)
                                            Else
                                                currentWS = CType(appInstance.Worksheets(sprintEndConfig.sheetDescript), Global.Microsoft.Office.Interop.Excel.Worksheet)
                                            End If
                                        End If
                                        With currentWS
                                            Select Case sprintEndConfig.Typ
                                                Case "Text"
                                                    oneNextTask.SprintEndDate = CStr(.Cells(i, sprintEndConfig.column.von).value)
                                                'Case "Integer"
                                                '    oneNextTask.Sprint.SprintEndDate = CInt(.Cells(i, sprintEndConfig.column.von).value)
                                                'Case "Decimal"
                                                '    oneNextTask.Sprint.SprintEndDate = CDbl(.Cells(i, sprintEndConfig.column.von).value)
                                                Case "Date"
                                                    oneNextTask.SprintEndDate = CDate(.Cells(i, sprintEndConfig.column.von).value)
                                                Case Else
                                                    oneNextTask.SprintEndDate = .Cells(i, sprintEndConfig.column.von).value
                                            End Select

                                            If sprintEndConfig.objType = "RegEx" Then
                                                regexpression = New Regex(sprintEndConfig.content)
                                                Dim match As Match = regexpression.Match(oneNextTask.SprintEndDate)
                                                If match.Success Then
                                                    oneNextTask.SprintEndDate = match.Value
                                                Else
                                                    oneNextTask.SprintEndDate = Nothing
                                                End If
                                            End If
                                        End With

                                    Catch ex As Exception
                                        ' Fehler bei  SprintEndDate
                                        oneNextTask.SprintEndDate = Nothing
                                        outputline = "problems reading SprintEndDate: line: " & i.ToString
                                        'meldungen.Add(outputline)
                                        Call logger(ptErrLevel.logError, outputline, "readJIRATasks", anzFehler)
                                    End Try

                                    ' SprintCompleteDate:
                                    Try
                                        Dim sprintCompleteConfig As clsConfigProjectsImport = projectConfig("SprintCompleteDate")
                                        'richtige Tabelle öffnen

                                        If currentWS.Index <> sprintCompleteConfig.sheet Then
                                            If Not IsNothing(sprintCompleteConfig.sheet) Then
                                                currentWS = CType(appInstance.Worksheets(sprintCompleteConfig.sheet), Global.Microsoft.Office.Interop.Excel.Worksheet)
                                            Else
                                                currentWS = CType(appInstance.Worksheets(sprintCompleteConfig.sheetDescript), Global.Microsoft.Office.Interop.Excel.Worksheet)
                                            End If
                                        End If
                                        With currentWS
                                            Dim sprintCompleteObject As Object = Nothing
                                            Try
                                                sprintCompleteObject = .Cells(i, sprintCompleteConfig.column.von).value
                                                If sprintCompleteObject = "[no subfield found]" Then
                                                    sprintCompleteObject = Nothing
                                                End If
                                            Catch ex As Exception
                                                sprintCompleteObject = Nothing
                                            End Try


                                            If Not IsNothing(sprintCompleteObject) Then
                                                Select Case sprintCompleteConfig.Typ
                                                    Case "Text"
                                                        oneNextTask.SprintCompleteDate = CStr(.Cells(i, sprintCompleteConfig.column.von).value)
                                                'Case "Integer"
                                                '    oneNextTask.SprintCompleteDate = CInt(.Cells(i, sprintEndConfig.column.von).value)
                                                'Case "Decimal"
                                                '    oneNextTask.SprintCompleteDate = CDbl(.Cells(i, sprintEndConfig.column.von).value)
                                                    Case "Date"
                                                        oneNextTask.SprintCompleteDate = CDate(.Cells(i, sprintCompleteConfig.column.von).value)
                                                    Case Else
                                                        oneNextTask.SprintCompleteDate = .Cells(i, sprintCompleteConfig.column.von).value
                                                End Select

                                                If sprintCompleteConfig.objType = "RegEx" Then
                                                    regexpression = New Regex(sprintCompleteConfig.content)
                                                    Dim match As Match = regexpression.Match(oneNextTask.SprintCompleteDate)
                                                    If match.Success Then
                                                        oneNextTask.SprintCompleteDate = match.Value
                                                    Else
                                                        oneNextTask.SprintCompleteDate = Nothing
                                                    End If
                                                End If
                                            Else
                                                oneNextTask.SprintCompleteDate = Date.MaxValue
                                            End If

                                        End With

                                    Catch ex As Exception
                                        ' Fehler bei  SprintCompleteDate
                                        outputline = "problems reading SprintCompleteDate: line: " & i.ToString
                                        meldungen.Add(outputline)
                                        Call logger(ptErrLevel.logError, outputline, "readJIRATasks", anzFehler)
                                    End Try

                                    ' SprintGoal:
                                    Try
                                        Dim sprintGoalConfig As clsConfigProjectsImport = projectConfig("SprintGoal")
                                        'richtige Tabelle öffnen
                                        If currentWS.Index <> sprintGoalConfig.sheet Then
                                            If Not IsNothing(sprintGoalConfig.sheet) Then
                                                currentWS = CType(appInstance.Worksheets(sprintGoalConfig.sheet), Global.Microsoft.Office.Interop.Excel.Worksheet)
                                            Else
                                                currentWS = CType(appInstance.Worksheets(sprintGoalConfig.sheetDescript), Global.Microsoft.Office.Interop.Excel.Worksheet)
                                            End If
                                        End If
                                        With currentWS
                                            Select Case sprintGoalConfig.Typ
                                                Case "Text"
                                                    oneNextTask.SprintGoal = CStr(.Cells(i, sprintGoalConfig.column.von).value)
                                                Case "Integer"
                                                    oneNextTask.SprintGoal = CInt(.Cells(i, sprintGoalConfig.column.von).value)
                                                Case "Decimal"
                                                    oneNextTask.SprintGoal = CDbl(.Cells(i, sprintGoalConfig.column.von).value)
                                                Case "Date"
                                                    oneNextTask.SprintGoal = CDate(.Cells(i, sprintGoalConfig.column.von).value)
                                                Case Else
                                                    oneNextTask.SprintGoal = .Cells(i, sprintGoalConfig.column.von).value
                                            End Select

                                            If sprintGoalConfig.objType = "RegEx" Then
                                                regexpression = New Regex(sprintGoalConfig.content)
                                                Dim match As Match = regexpression.Match(oneNextTask.SprintGoal)
                                                If match.Success Then
                                                    oneNextTask.SprintGoal = match.Value
                                                Else
                                                    oneNextTask.SprintGoal = Nothing
                                                End If
                                            End If
                                        End With

                                    Catch ex As Exception
                                        ' Fehler bei  SprintGoal
                                        outputline = "problems reading SprintGoal: line: " & i.ToString
                                        meldungen.Add(outputline)
                                        Call logger(ptErrLevel.logError, outputline, "readJIRATasks", anzFehler)
                                    End Try

                                    '------------------------------------------
                                    ' Plausibilitäten prüfen und ggfs. anpassen
                                    '------------------------------------------
                                    '' wenn kein Fälligkeitdatum angegeben ist, so wird date.maxvalue hierfür gesetzt
                                    'If oneNextTask.fällig <= Date.MinValue Then
                                    '    oneNextTask.fällig = Date.MaxValue
                                    'End If
                                    ' wenn task einem Sprint zugeordnet, so ist Fälligkeit = enddate des Sprints
                                    If Not IsNothing(oneNextTask.SprintName) And Not IsNothing(oneNextTask.SprintEndDate) Then
                                        oneNextTask.fällig = oneNextTask.SprintEndDate
                                    End If

                                    ' ist die Task bereits fertig ,
                                    If oneNextTask.TaskStatus = "Fertig" Then
                                        oneNextTask.fällig = oneNextTask.erledigt
                                    End If
                                    '?????UR: TODO
                                    If Not projListSorted.ContainsKey(oneNextTask.projectName) Then
                                        taskListSorted = New SortedList(Of String, clsJIRA_Task)
                                    Else
                                        taskListSorted = projListSorted(oneNextTask.projectName)
                                    End If

                                    If Not taskListSorted.ContainsKey(oneNextTask.Jira_ID) Then
                                        taskListSorted.Add(oneNextTask.Jira_ID, oneNextTask)
                                    Else ' sollte eigentlich nicht vorkommen
                                        taskListSorted.Remove(oneNextTask.Jira_ID)
                                        taskListSorted.Add(oneNextTask.Jira_ID, oneNextTask)
                                    End If

                                    If Not projListSorted.ContainsKey(oneNextTask.projectName) Then
                                        projListSorted.Add(oneNextTask.projectName, taskListSorted)
                                    Else
                                        projListSorted.Remove(oneNextTask.projectName)
                                        projListSorted.Add(oneNextTask.projectName, taskListSorted)
                                    End If

                                    If Not projtaskList.ContainsKey(oneNextTask.projectName) Then
                                        taskList = New SortedList(Of Date, clsJIRA_Task)
                                    Else
                                        taskList = projtaskList(oneNextTask.projectName)
                                    End If

                                    While taskList.ContainsKey(oneNextTask.Erstellt)
                                        oneNextTask.Erstellt = DateAdd(DateInterval.Second, 1, oneNextTask.Erstellt)
                                    End While
                                    If Not IsNothing(oneNextTask.Erstellt) Then
                                        taskList.Add(oneNextTask.Erstellt, oneNextTask)
                                    End If

                                    If Not projtaskList.ContainsKey(oneNextTask.projectName) Then
                                        projtaskList.Add(oneNextTask.projectName, taskList)
                                    Else
                                        projtaskList.Remove(oneNextTask.projectName)
                                        projtaskList.Add(oneNextTask.projectName, taskList)
                                    End If

                                    If Not projsprintList.ContainsKey(oneNextTask.projectName) Then
                                        sprintList = New SortedList(Of String, clsJIRA_sprint)
                                    Else
                                        sprintList = projsprintList(oneNextTask.projectName)
                                    End If

                                    ' SprintList erstellen
                                    Dim sprintItem As clsJIRA_sprint
                                    If Not IsNothing(oneNextTask.SprintName) Then
                                        If Not sprintList.ContainsKey(oneNextTask.SprintName) Then
                                            sprintItem = New clsJIRA_sprint
                                            sprintItem.SprintName = oneNextTask.SprintName
                                            sprintItem.SprintStartDate = oneNextTask.SprintStartDate
                                            sprintItem.SprintEndDate = oneNextTask.SprintEndDate
                                            sprintItem.SprintCompleteDate = oneNextTask.SprintCompleteDate
                                            sprintItem.SprintGoal = oneNextTask.SprintGoal
                                            sprintItem.SprintTasks.Add(oneNextTask.Jira_ID, oneNextTask.Jira_ID)
                                            sprintList.Add(sprintItem.SprintName, sprintItem)
                                        Else
                                            sprintItem = New clsJIRA_sprint
                                            sprintItem = sprintList.Item(oneNextTask.SprintName)
                                            sprintItem.SprintTasks.Add(oneNextTask.Jira_ID, oneNextTask.Jira_ID)
                                            sprintList.Remove(oneNextTask.SprintName)
                                            sprintList.Add(sprintItem.SprintName, sprintItem)
                                        End If
                                    End If
                                    If Not projsprintList.ContainsKey(oneNextTask.projectName) Then
                                        projsprintList.Add(oneNextTask.projectName, sprintList)
                                    Else
                                        projsprintList.Remove(oneNextTask.projectName)
                                        projsprintList.Add(oneNextTask.projectName, sprintList)
                                    End If

                                    outputline = "line " & i.ToString & " Done"
                                    Call logger(ptErrLevel.logInfo, outputline, "readJIRATasks", anzFehler)

                                Next i       ' next line

                            Catch ex As Exception
                                outputline = "problems reading project: catch ex of lines-loop : " & ex.Message
                                meldungen.Add(outputline)
                                Call MsgBox(outputline)
                            End Try

                        End If      ' worksheet gibt es nicht

                    End If


                Catch ex As Exception
                    outputline = "problems reading project: catch ex of projectWB : " & ex.Message
                    meldungen.Add(outputline)
                    Call MsgBox(outputline)

                End Try
            End If

            ' Schliessen der Excel-Ausleitung von Jira


            appInstance.Workbooks(projectWB.Name).Close(SaveChanges:=True)

        Catch ex As Exception
            outputline = "problems reading project: catch from the beginning of readJIRATasks : " & ex.Message
            meldungen.Add(outputline)
            Call MsgBox(outputline)
        End Try

        If meldungen.Count <= 0 Then
            result = True
        End If

        readJIRATasks = result

    End Function



    Public Sub writeYearInitialPlanningSupportToExcel(ByVal von As Integer, ByVal bis As Integer,
                                                      ByVal roleCollection As Collection, ByVal costCollection As Collection,
                                                      ByVal unit As PTEinheiten)
        appInstance.EnableEvents = False

        ' wenn CostCollection was enthält , dann wird unit automatisch auf Euro gesetzt 
        ' andernfalls wäre die Ressourcen in PT, die Kostenarten Zahlen in T€ zu interpretieren und das ist strange 
        If Not IsNothing(costCollection) Then
            If costCollection.Count > 0 Then
                unit = PTEinheiten.euro
            End If
        End If


        Dim projectsToWork As New Collection
        Dim defDone As Boolean = False
        If Not IsNothing(selectedProjekte) Then
            If selectedProjekte.Count > 0 Then
                For Each kvp As KeyValuePair(Of String, clsProjekt) In selectedProjekte.Liste
                    If Not projectsToWork.Contains(kvp.Key) Then
                        projectsToWork.Add(kvp.Key, kvp.Key)
                    End If
                Next
                defDone = True
            End If
        End If


        If Not defDone And ShowProjekte.getMarkedProjects.Count > 0 Then
            projectsToWork = ShowProjekte.getMarkedProjects
            defDone = True
        End If

        If Not defDone Then
            For Each kvp As KeyValuePair(Of String, clsProjekt) In ShowProjekte.Liste
                projectsToWork.Add(kvp.Key, kvp.Key)
            Next
        End If


        Dim newWB As Excel.Workbook

        Dim considerAll As Boolean = True



        Dim fNameExtension As String = ""
        ' den Dateinamen bestimmen ...


        Dim expFName As String = exportOrdnerNames(PTImpExp.massenEdit) & "\Soll-Ist-Kapa Planning " & fNameExtension & ".xlsx"


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

        Dim startSpalteDaten As Integer = 8
        Dim roleCostNames As Excel.Range = Nothing
        Dim roleCostInput As Excel.Range = Nothing

        Dim tmpName As String = ""


        With CType(newWB.Worksheets("VISBO"), Excel.Worksheet)
            Dim ersteZeile As Excel.Range
            ersteZeile = CType(.Range(.Cells(1, 1), .Cells(1, 6 + bis - von)), Excel.Range)

            CType(.Cells(1, 1), Excel.Range).Value = "Project-Name"
            CType(.Cells(1, 2), Excel.Range).Value = "Project-Number"
            CType(.Cells(1, 3), Excel.Range).Value = "Variant-Name"
            CType(.Cells(1, 4), Excel.Range).Value = "Version"
            CType(.Cells(1, 5), Excel.Range).Value = "Reference-Date"

            If unit = PTEinheiten.euro Then
                CType(.Cells(1, 6), Excel.Range).Value = "Ressource-/Cost-Name"
            ElseIf unit = PTEinheiten.hrs Then
                CType(.Cells(1, 6), Excel.Range).Value = "Ressource-Name"
            ElseIf unit = PTEinheiten.personentage Then
                CType(.Cells(1, 6), Excel.Range).Value = "Ressource-Name"
            Else
                CType(.Cells(1, 6), Excel.Range).Value = "Ressource-/Cost-Name"
            End If

            CType(.Cells(1, 7), Excel.Range).Value = "Type"

            ' damit das beim programmatischen auslesen auch berücksichtigt werden kann 
            CType(.Cells(1, 7), Excel.Range).ClearComments()
            CType(.Cells(1, 7), Excel.Range).AddComment(unit.ToString)

            ' jetzt wird die Zeile 1 geschrieben 
            Dim startMonat As Date = StartofCalendar.AddMonths(von - 1)


            ' jetzt werden die Überschriften des Datenbereichs geschrieben 
            For m As Integer = 0 To bis - von
                With CType(.Cells(1, startSpalteDaten + m), Global.Microsoft.Office.Interop.Excel.Range)
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
            Next


        End With

        Dim ws As Excel.Worksheet = CType(newWB.Worksheets("VISBO"), Excel.Worksheet)

        zeile = 2
        Dim zeitraum As Integer = bis - von

        Dim lastplanProjekte As New clsProjekte
        Dim beauftragungsProjekte As New clsProjekte

        Dim lastDate As Date = Date.Now.AddMonths(-1).AddDays(24 - Date.Now.Day)
        Dim heute As Date = Date.Now

        Dim err As New clsErrorCodeMsg

        If Not IsNothing(roleCollection) Then

            For i As Integer = 1 To roleCollection.Count

                Dim roleNameID As String = roleCollection.Item(i)
                Dim teamID As Integer = -1
                Dim curRole As clsRollenDefinition = RoleDefinitions.getRoleDefByIDKennung(roleNameID, teamID)

                Dim myCollection As New Collection From {curRole.name}

                Dim kapaValues() As Double = ShowProjekte.getRoleKapasInMonth(myCollection)
                Call writePlanningDataRow(newWB.Name, ws.Name, zeile, startSpalteDaten, Nothing,
                                                  von, bis, curRole, Nothing, unit, PTVergleichsArt.capacity, heute, kapaValues)

                zeile = zeile + 1


                For Each kvp As KeyValuePair(Of String, clsProjekt) In ShowProjekte.Liste

                    Dim vpID As String = ""
                    Dim lastplan As clsProjekt = getProjektFromSessionOrDB(kvp.Value.name, kvp.Value.variantName, AlleProjekte, lastDate)

                    Dim lastPlanValues() As Double = Nothing
                    If Not IsNothing(lastplan) Then
                        ' jetzt die Werte für den letzten Plan schreiben 
                        lastPlanValues = lastplan.getResourceValuesInTimeFrame(von, bis, roleNameID, inclSubRoles:=True, outPutInEuro:=False)
                        Call writePlanningDataRow(newWB.Name, ws.Name, zeile, startSpalteDaten, lastplan,
                                                  von, bis, curRole, Nothing, unit, PTVergleichsArt.formerPlan, lastDate, lastPlanValues)


                        If Not lastplanProjekte.contains(lastplan.name) Then
                            lastplanProjekte.Add(lastplan, False)
                        End If

                        zeile = zeile + 1
                    End If

                    Dim beauftragung As clsProjekt = getProjektFromSessionOrDB(kvp.Value.name, ptVariantFixNames.pfv.ToString, AlleProjekte, heute)
                    Dim a As Integer = beauftragung.dauerInDays
                    Dim baselineValues() As Double = Nothing

                    If Not IsNothing(beauftragung) Then
                        ' jetzt die Werte für die Beauftragung schreiben 
                        baselineValues = beauftragung.getResourceValuesInTimeFrame(von, bis, roleNameID, inclSubRoles:=True, outPutInEuro:=False)
                        Call writePlanningDataRow(newWB.Name, ws.Name, zeile, startSpalteDaten, beauftragung,
                                                  von, bis, curRole, Nothing, unit, PTVergleichsArt.beauftragung, heute, baselineValues)

                        If Not beauftragungsProjekte.contains(beauftragung.name) Then
                            beauftragungsProjekte.Add(beauftragung, False)
                        End If

                        zeile = zeile + 1
                    End If

                    Dim planValues() As Double = kvp.Value.getResourceValuesInTimeFrame(von, bis, roleNameID, inclSubRoles:=True, outPutInEuro:=False)
                    Call writePlanningDataRow(newWB.Name, ws.Name, zeile, startSpalteDaten, kvp.Value,
                                                  von, bis, curRole, Nothing, unit, PTVergleichsArt.planungsstand, heute, planValues)

                    zeile = zeile + 1

                Next

                ' now print the sums of it ..
                Dim sumValues() As Double = lastplanProjekte.getRoleValuesInMonth(roleIDStr:=roleNameID, considerAllSubRoles:=True, considerAllNeedsOfRolesHavingTheseSkills:=True)
                Call writePlanningDataRow(newWB.Name, ws.Name, zeile, startSpalteDaten, Nothing,
                                                  von, bis, curRole, Nothing, unit, PTVergleichsArt.formerPlan, lastDate, sumValues)

                zeile = zeile + 1

                sumValues = ShowProjekte.getRoleValuesInMonth(roleIDStr:=roleNameID, considerAllSubRoles:=True, considerAllNeedsOfRolesHavingTheseSkills:=True)
                Call writePlanningDataRow(newWB.Name, ws.Name, zeile, startSpalteDaten, Nothing,
                                                  von, bis, curRole, Nothing, unit, PTVergleichsArt.planungsstand, heute, sumValues)

                zeile = zeile + 1

                sumValues = beauftragungsProjekte.getRoleValuesInMonth(roleIDStr:=roleNameID, considerAllSubRoles:=True, considerAllNeedsOfRolesHavingTheseSkills:=True)
                Call writePlanningDataRow(newWB.Name, ws.Name, zeile, startSpalteDaten, Nothing,
                                                  von, bis, curRole, Nothing, unit, PTVergleichsArt.beauftragung, heute, sumValues)

                zeile = zeile + 1

            Next

        End If

        If Not IsNothing(costCollection) Then


            For i As Integer = 1 To costCollection.Count


                Dim curCost As clsKostenartDefinition = CostDefinitions.getCostdef(costCollection.Item(i))

                For Each kvp As KeyValuePair(Of String, clsProjekt) In ShowProjekte.Liste

                    Dim lastplan As clsProjekt = getProjektFromSessionOrDB(kvp.Value.name, kvp.Value.variantName, AlleProjekte, lastDate)
                    Dim lastPlanValues() As Double = Nothing
                    If Not IsNothing(lastplan) Then
                        ' jetzt die Werte für die Beauftragung schreiben 
                        lastPlanValues = lastplan.getCostValuesInTimeFrame(von, bis, curCost.name)
                        Call writePlanningDataRow(newWB.Name, ws.Name, zeile, startSpalteDaten, lastplan,
                                                  von, bis, Nothing, curCost, PTEinheiten.euro, PTVergleichsArt.planungsstand, lastDate, lastPlanValues)

                        zeile = zeile + 1
                    End If


                    Dim beauftragung As clsProjekt = getProjektFromSessionOrDB(kvp.Value.name, ptVariantFixNames.pfv.ToString, AlleProjekte, Date.Now)
                    Dim baselineValues() As Double = Nothing

                    If Not IsNothing(beauftragung) Then
                        ' jetzt die Werte für die Beauftragung schreiben 
                        baselineValues = beauftragung.getCostValuesInTimeFrame(von, bis, curCost.name)
                        Call writePlanningDataRow(newWB.Name, ws.Name, zeile, startSpalteDaten, beauftragung,
                                                  von, bis, Nothing, curCost, PTEinheiten.euro, PTVergleichsArt.beauftragung, heute, baselineValues)

                        zeile = zeile + 1
                    End If

                    Dim bedarfsValues() As Double = kvp.Value.getCostValuesInTimeFrame(von, bis, curCost.name)

                    Call writePlanningDataRow(newWB.Name, ws.Name, zeile, startSpalteDaten, kvp.Value,
                                                  von, bis, Nothing, curCost, PTEinheiten.euro, PTVergleichsArt.planungsstand, heute, bedarfsValues)

                    zeile = zeile + 1

                Next

                ' now print the sums of it ..
                Dim sumValues() As Double = lastplanProjekte.getCostValuesInMonthNew(curCost.name)
                Call writePlanningDataRow(newWB.Name, ws.Name, zeile, startSpalteDaten, Nothing,
                                                  von, bis, Nothing, curCost, PTEinheiten.euro, PTVergleichsArt.formerPlan, lastDate, sumValues)

                zeile = zeile + 1

                sumValues = ShowProjekte.getCostValuesInMonthNew(curCost.name)
                Call writePlanningDataRow(newWB.Name, ws.Name, zeile, startSpalteDaten, Nothing,
                                                  von, bis, Nothing, curCost, PTEinheiten.euro, PTVergleichsArt.planungsstand, heute, sumValues)

                zeile = zeile + 1

                sumValues = beauftragungsProjekte.getCostValuesInMonthNew(curCost.name)
                Call writePlanningDataRow(newWB.Name, ws.Name, zeile, startSpalteDaten, Nothing,
                                                  von, bis, Nothing, curCost, PTEinheiten.euro, PTVergleichsArt.beauftragung, heute, sumValues)

                zeile = zeile + 1

            Next

        End If

        ' jetzt werden die Summen über alle Rollen und Kosten gebildet ...
        ' siehe kapavalues ... 


        Try
            ' jetzt die Autofilter aktivieren ... 
            If Not CType(newWB.Worksheets("VISBO"), Excel.Worksheet).AutoFilterMode = True Then
                CType(newWB.Worksheets("VISBO"), Excel.Worksheet).Cells(1, 1).AutoFilter()
            End If

            newWB.Close(SaveChanges:=True)
        Catch ex As Exception
            Throw New ArgumentException("Fehler beim Speichern" & ex.Message)
        End Try

        appInstance.EnableEvents = True

        Call MsgBox("ok, Datei exportiert")

    End Sub

    ''' <summary>
    ''' schreibt im Export von writeYearInitialPlanningSupport eine Zeile von Kapazität, Planung oder Beauftragungs-Wert
    ''' </summary>
    ''' <param name="wbName"></param>
    ''' <param name="wsName"></param>
    ''' <param name="zeile"></param>
    ''' <param name="startSpalteDaten"></param>
    ''' <param name="hproj"></param>
    ''' <param name="von"></param>
    ''' <param name="bis"></param>
    ''' <param name="curRole"></param>
    ''' <param name="curCost"></param>
    ''' <param name="unit"></param>
    ''' <param name="vglType"></param>
    ''' <param name="values"></param>
    Private Sub writePlanningDataRow(ByVal wbName As String, ByVal wsName As String, ByVal zeile As Integer, ByVal startSpalteDaten As Integer, ByVal hproj As clsProjekt,
                                     ByVal von As Integer, ByVal bis As Integer, ByVal curRole As clsRollenDefinition, ByVal curCost As clsKostenartDefinition,
                                     ByVal unit As PTEinheiten, ByVal vglType As PTVergleichsArt, ByVal vglDate As Date, ByVal values As Double())

        Dim typeStrings As String() = {"Kapazität", "Beauftragung", "Planung (akt)", "Planung (vom)"}
        If awinSettings.englishLanguage Then
            typeStrings = {"Capacity", "Baseline", "Planning (current)", "Planning (former)"}
        End If

        Dim actualDataIndex As Integer = -1
        Try
            If Not IsNothing(hproj) Then
                If hproj.hasActualValues Then
                    actualDataIndex = getColumnOfDate(hproj.actualDataUntil) - getColumnOfDate(hproj.startDate)
                End If
            End If
        Catch ex As Exception

        End Try


        Dim formatierung As String = "#,##0.##"
        Dim typBezeichner As String = ""

        If vglType = PTVergleichsArt.planungsstand Then
            typBezeichner = typeStrings(2)
        ElseIf vglType = PTVergleichsArt.beauftragung Then
            typBezeichner = typeStrings(1)
        ElseIf vglType = PTVergleichsArt.capacity Then
            typBezeichner = typeStrings(0)
        ElseIf vglType = PTVergleichsArt.formerPlan Then
            typBezeichner = typeStrings(3)
        End If

        Dim ws As Excel.Worksheet = CType(appInstance.Workbooks.Item(wbName).Worksheets(wsName), Excel.Worksheet)

        ' wenn es eine Rolle ist, müssen die Values ggf umgerechnet werden ... 
        If Not IsNothing(curRole) Then
            If unit = PTEinheiten.euro Then
                For ix As Integer = 0 To values.Length - 1
                    values(ix) = values(ix) * curRole.tagessatzIntern
                Next
            ElseIf unit = PTEinheiten.hrs Then
                For ix As Integer = 0 To values.Length - 1
                    values(ix) = values(ix) * 8
                Next
            End If
        End If


        If Not vglType = PTVergleichsArt.none And Not IsNothing(hproj) Then
            CType(ws.Cells(zeile, 1), Excel.Range).Value = hproj.name
            CType(ws.Cells(zeile, 2), Excel.Range).Value = hproj.kundenNummer
            CType(ws.Cells(zeile, 3), Excel.Range).Value = hproj.variantName
            CType(ws.Cells(zeile, 4), Excel.Range).Value = hproj.timeStamp.ToShortDateString
            CType(ws.Cells(zeile, 5), Excel.Range).Value = vglDate.ToShortDateString
        ElseIf IsNothing(hproj) Then
            CType(ws.Cells(zeile, 1), Excel.Range).Value = getPnameFromKey(currentConstellationPvName)
            CType(ws.Cells(zeile, 2), Excel.Range).Value = ""
            CType(ws.Cells(zeile, 3), Excel.Range).Value = getVariantnameFromKey(currentConstellationPvName)
            CType(ws.Cells(zeile, 4), Excel.Range).Value = ""
            CType(ws.Cells(zeile, 5), Excel.Range).Value = vglDate.ToShortDateString
        End If

        If Not IsNothing(curRole) Then
            CType(ws.Cells(zeile, 6), Excel.Range).Value = curRole.name
        ElseIf Not IsNothing(curCost) Then
            CType(ws.Cells(zeile, 6), Excel.Range).Value = curCost.name
        End If


        If unit = PTEinheiten.personentage Then
            CType(ws.Cells(zeile, 7), Excel.Range).Value = typBezeichner & " [PT]"
        ElseIf unit = PTEinheiten.euro Then
            CType(ws.Cells(zeile, 7), Excel.Range).Value = typBezeichner & " [T€]"
        ElseIf unit = PTEinheiten.hrs Then
            CType(ws.Cells(zeile, 7), Excel.Range).Value = typBezeichner & " [Hrs]"
        Else
            CType(ws.Cells(zeile, 7), Excel.Range).Value = typBezeichner & "[?]"
        End If

        Dim editRange As Excel.Range = CType(ws.Range(ws.Cells(zeile, startSpalteDaten), ws.Cells(zeile, startSpalteDaten + bis - von)), Excel.Range)
        editRange.Value = values
        editRange.NumberFormat = formatierung

        If actualDataIndex >= 0 Then
            CType(ws.Range(ws.Cells(zeile, startSpalteDaten), ws.Cells(zeile, startSpalteDaten + actualDataIndex)), Excel.Range).Interior.Color = RGB(227, 227, 227)
            CType(ws.Range(ws.Cells(zeile, startSpalteDaten), ws.Cells(zeile, startSpalteDaten + actualDataIndex)), Excel.Range).Locked = True
        End If


    End Sub



    ''' <summary>
    ''' legt mit den entsprechenden Angaben ein neues Projekt auf Basis der Vorlage an 
    ''' </summary>
    ''' <param name="pname"></param>
    ''' <param name="vorlagenName">Name der Projekt-Vorlage, die verwendet werden soll </param>
    ''' <param name="startdate">Startdatum</param>
    ''' <param name="endedate">Ende-Datum</param>
    ''' <param name="budgetVorgabe">Budget in T€, darf nicht negativ sein</param>
    ''' <param name="tafelZeile"></param>
    ''' <param name="sfit">Integer Wert für Strategie KPI, zwischen 1 und 9 </param>
    ''' <param name="risk">Integer Wert für Risiko KPI, zwischen 1 und 9</param>
    ''' <param name="profitUserAskedFor">Nothing oder positiver Wert zwischen 0 und 1, 0.1 heisst 10%</param>
    ''' <param name="kurzBeschreibung">Kurzbeschreibung des Projektes</param>
    ''' <param name="kdNr">Frei-Text für vom Kunden vergebene Projekt-Nummer; wird beim Input auf Zulässigkeit überprüft  </param>
    ''' <returns></returns>
    Public Function erstelleProjektAusVorlage(ByVal myproject As clsProjekt, ByVal pname As String, ByVal vorlagenName As String, ByVal startdate As Date,
                                ByVal endedate As Date, ByVal budgetVorgabe As Double,
                                ByVal tafelZeile As Integer, ByVal sfit As Double, ByVal risk As Double, ByVal profitUserAskedFor As String,
                                ByVal kurzBeschreibung As String, ByVal buName As String, Optional ByVal kdNr As String = "", Optional ByVal template As Boolean = False) As clsProjekt
        Dim newprojekt As Boolean
        Dim hproj As clsProjekt
        Dim pStatus As String = ProjektStatus(0)
        Dim zeile As Integer = tafelZeile
        'Dim spalte As Integer = start
        Dim plen As Integer
        'Dim top As Double, left As Double, width As Double, height As Double
        'Dim shpElement As Excel.Shape
        Dim pcolor As Object
        Dim heute As Date = Date.Now
        Dim heute1 As Date = Now
        Dim key As String = pname & "#"
        Dim ms As Long = heute.Millisecond
        Dim zielrenditenVorgabe As Double = Nothing
        Dim zielrenditenVorgabe1 As Double = Nothing
        Dim zielrenditenVorgabe2 As Double = Nothing
        newprojekt = True
        Dim err As New clsErrorCodeMsg
        '
        ' ein neues Projekt wird als Objekt angelegt ....
        '

        hproj = New clsProjekt


        If Projektvorlagen.Contains(vorlagenName) Or Not IsNothing(myproject) Then

            ' Aufruf von Rest-Call Create a Copy of a Version url: https://dev.visbo.net/api/vpv/:vpvid/copy
            If Not template And IsNothing(myproject) Then
                Try
                    hproj = CType(databaseAcc, DBAccLayer.Request).createProjectFromTemplate(pname, vorlagenName, startdate, endedate, budgetVorgabe, sfit, risk, profitUserAskedFor, kurzBeschreibung, kdNr, err)
                Catch ex As Exception
                    Call MsgBox("createProjectFromTemplate - konnte kein Projekt erzeugen, evt. das Template nicht in der DB")
                End Try
            Else
                hproj = Nothing
            End If


            If IsNothing(hproj) Then
                hproj = New clsProjekt


                ' jetzt wird bestimmt, ob es eine Zielrenditen Vorgabe gibt ... 
                If IsNothing(profitUserAskedFor) Then
                    ' nichts weiter tun ... zielrenditenVorgabe ist mit Nothing besetzt 
                Else
                    If IsNumeric(profitUserAskedFor) Then
                        'Dim referenceBudget As Double = Projektvorlagen.getProject(vorlagenName).getSummeKosten
                        Dim referenceBudget As Double
                        If Not IsNothing(myproject) Then
                            referenceBudget = myproject.Erloes
                        Else
                            referenceBudget = Projektvorlagen.getProject(vorlagenName).Erloes
                        End If

                        If referenceBudget > 0 Then
                            'Dim verfuegbaresBudget As Double = budgetVorgabe / (CDbl(profitUserAskedFor) / 100 + 1)
                            'zielrenditenVorgabe = verfuegbaresBudget / referenceBudget
                            'zielrenditenVorgabe1 = (budgetVorgabe * (CDbl(profitUserAskedFor) / 100 + 1)) / referenceBudget
                            zielrenditenVorgabe = (budgetVorgabe * (1 - CDbl(profitUserAskedFor) / 100)) / referenceBudget
                        End If

                    Else
                        Call MsgBox("keine zulässige Renditen Angabe ...")
                        erstelleProjektAusVorlage = Nothing
                        Exit Function
                    End If
                End If


                Try
                    ' Projektdauer wurde durch Start- und Endedatum im Formular angegeben
                    If Not IsNothing(myproject) Then
                        Dim projVorlage As New clsProjektvorlage
                        projVorlage.VorlagenName = myproject.name
                        projVorlage.Schrift = myproject.Schrift
                        projVorlage.Schriftfarbe = myproject.Schriftfarbe
                        'projVorlage.farbe = myproject.farbe
                        projVorlage.earliestStart = -6
                        projVorlage.latestStart = 6
                        projVorlage.Erloes = myproject.Erloes
                        projVorlage.AllPhases = myproject.AllPhases
                        projVorlage.hierarchy = myproject.hierarchy

                        projVorlage.korrCopyTo(hproj, startdate, endedate, zielrenditenVorgabe)
                    Else
                        Projektvorlagen.getProject(vorlagenName).korrCopyTo(hproj, startdate, endedate, zielrenditenVorgabe)
                    End If


                Catch ex As Exception
                    Call MsgBox("es gibt keine entsprechende Vorlage ..")
                    erstelleProjektAusVorlage = Nothing
                    Exit Function
                End Try


                Try
                    With hproj
                        .name = pname
                        .VorlagenName = vorlagenName
                        .startDate = startdate
                        .kundenNummer = kdNr
                        .businessUnit = buName
                        .Erloes = budgetVorgabe
                        .earliestStartDate = .startDate.AddMonths(.earliestStart)
                        .latestStartDate = .startDate.AddMonths(.latestStart)
                        .Status = ProjektStatus(PTProjektStati.geplant)
                        .description = kurzBeschreibung

                        .StrategicFit = sfit
                        .Risiko = risk
                        plen = .anzahlRasterElemente
                        pcolor = .farbe
                    End With


                    ' nächste Zeile ist ein work-around für Fehler Der Index liegt außerhalb der Array-Grenzen
                    ' workaround
                    Dim tmpdata As Integer = hproj.dauerInDays
                    'Call awinCreateBudgetWerte(hproj)

                Catch ex As Exception
                    Call MsgBox(ex.Message)
                    erstelleProjektAusVorlage = Nothing
                    Exit Function
                End Try

                ' Anpassen der Daten für die Termine 
                ' wenn Samstag oder Sonntag, dann auf den Freitag davor legen 
                ' nein - das darf nicht gemacht werden; evtl liegt ja dann der Meilenstein vor der Phase 
                ' grundsätzlich sollte der Anwender hier bestimmen, nicht das Programm
            Else

                'TODO: update einer Vorlage muss auch funktionieren

            End If

        Else
            Call MsgBox("es gibt keine entsprechende Vorlage: " & vorlagenName)
            erstelleProjektAusVorlage = Nothing
            Exit Function
        End If

        '
        ' Ende Objekt Anlage
        erstelleProjektAusVorlage = hproj

    End Function



    '
    ' Sub trägt ein individuelles Projekt ein
    '
    ''' <summary>
    ''' trägt ein neues Projekt in Showprojekte ein
    ''' wenn myProject nicht Nothing ist, dann wird das als Vorlage hergenommen 
    ''' </summary>
    ''' <param name="pname">Projektname</param>
    ''' <param name="vorlagenName">Vorlagen-Name</param>
    ''' <param name="startdate">Start-Datum des PRojekts</param>
    ''' <param name="budgetVorgabe">Budget des Projekts</param>
    ''' <param name="tafelZeile">
    ''' in welcher Zeile der Projekt-Tafel soll es gezeichnet werden; 
    ''' 0:= finde eine geeignete Stelle
    ''' </param>
    ''' <param name="sfit">Wert für den strategischen Fit</param>
    ''' <param name="risk">Wert für das Risiko</param>
    ''' <param name="profitUserAskedFor">der Ergebnis Forecast in Prozent der Gesamtkosten, den der Nutzer gerne sehen möchte</param>
    ''' <remarks></remarks>
    Public Sub TrageivProjektein(ByVal myProject As clsProjekt,
                                ByVal pname As String, ByVal vorlagenName As String, ByVal startdate As Date,
                                 ByVal endedate As Date, ByVal budgetVorgabe As Double,
                                 ByVal tafelZeile As Integer, ByVal sfit As Double, ByVal risk As Double, ByVal profitUserAskedFor As String,
                                 ByVal kurzBeschreibung As String, ByVal kdNummer As String)

        Dim hproj As clsProjekt = Nothing
        Dim err As New clsErrorCodeMsg

        hproj = erstelleProjektAusVorlage(myProject, pname, vorlagenName, startdate, endedate, budgetVorgabe,
                                  tafelZeile, sfit, risk, profitUserAskedFor,
                                  kurzBeschreibung, "", kdNr:=kdNummer, template:=False)


        ' jetzt wird testweise das hproj.setMilestone Invoices gemacht ..
        'If hproj.name.StartsWith("E_Kunde") Then
        '    Call hproj.setMilestoneInvoices("Finalization")
        'End If

        Dim formerEE As Boolean = appInstance.EnableEvents
        Dim formerSU As Boolean = appInstance.ScreenUpdating
        appInstance.EnableEvents = False
        appInstance.ScreenUpdating = False

        If Not AlleProjekte.hasAnyConflictsWith(calcProjektKey(hproj.name, hproj.variantName), False) Then

            Try
                AlleProjekte.Add(hproj)
            Catch ex As Exception

            End Try

            Try
                ShowProjekte.Add(hproj)
                ' hier ist im hproj das attribut shpUID gesetzt , deswegen muss nicht extra AddShape aufgerufen werden 
            Catch ex As Exception

            End Try

            ' wenn bestimmte Projekte beim Suchen nach einem Platz nicht berücksichtigt werden sollen,
            ' dann müssen sie in einer Collection an ZeichneProjektinPlanTafel übergeben werden 
            Dim tmpCollection As New Collection

            '
            'Call awinClearPlanTafel()
            'Call awinZeichnePlanTafel(True)

            Dim pZeile As Integer = projectboardShapes.getMaxZeile
            Call ZeichneProjektinPlanTafel(tmpCollection, pname, pZeile, tmpCollection, tmpCollection)



            ' ein Projekt wurde eingefügt  - typus = 2
            Call awinNeuZeichnenDiagramme(2)

        Else
            Call MsgBox("Konflikte mit Summary Projekten")
        End If



        ' Call diagramsVisible("True")
        appInstance.ScreenUpdating = formerSU
        appInstance.EnableEvents = formerEE

    End Sub
End Module
