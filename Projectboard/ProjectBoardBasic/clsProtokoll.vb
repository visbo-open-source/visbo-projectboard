Imports ProjectBoardDefinitions
Imports xlNS = Microsoft.Office.Interop.Excel
Imports Microsoft.Office.Interop.Excel

Public Class clsProtokoll
    Private _tabblattname As String
    Private _actDate As String
    Private _Projekt As String
    Private _hierarchie As String
    Private _planelement As String
    Private _klasse As String
    Private _abkürzung As String
    Private _quelle As String
    Private _planeleÜbern As String
    Private _grund As String
    Private _PThierarchie As String
    Private _PTklasse As String

    ''' <summary>
    ''' Liest und schreibt den Namen des Tabellenblattes
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Property tabblattname As String
        Get
            tabblattname = _tabblattname
        End Get
        Set(value As String)
            _tabblattname = value
        End Set
    End Property
    ''' <summary>
    ''' Liest und schreibt Datum der Aktion im Logbuch
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Property actDate As String
        Get
            actDate = _actDate
        End Get
        Set(value As String)
            _actDate = value
        End Set
    End Property
    ''' <summary>
    ''' Liest und schreibt das Projekt
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Property Projekt As String
        Get
            Projekt = _Projekt
        End Get
        Set(value As String)
            _Projekt = value
        End Set
    End Property
    ''' <summary>
    ''' Liest und schreibt ProjektHierarchie
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Property hierarchie As String
        Get
            hierarchie = _hierarchie
        End Get
        Set(value As String)
            _hierarchie = value
        End Set
    End Property

    ''' <summary>
    ''' Liest und schreibt  Import-Datei-Planelement
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Property planelement As String
        Get
            planelement = _planelement
        End Get
        Set(value As String)
            _planelement = value
        End Set
    End Property

    ''' <summary>
    ''' Liest und schreibt Klasse in Import-datei
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Property klasse As String
        Get
            klasse = _klasse
        End Get
        Set(value As String)
            _klasse = value
        End Set
    End Property

    ''' <summary>
    ''' Liest und schreibt Abkürzung des Planelements
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Property abkürzung As String
        Get
            abkürzung = _abkürzung
        End Get
        Set(value As String)
            _abkürzung = value
        End Set
    End Property
    ''' <summary>
    ''' Liest und schreibt Quelle (ImportDatei)
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Property quelle As String
        Get
            quelle = _quelle
        End Get
        Set(value As String)
            _quelle = value
        End Set
    End Property
    ''' <summary>
    ''' Liest und schreibt Name der Übernahme dieses Planelements ins Projekt
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Property planeleÜbern As String
        Get
            planeleÜbern = _planeleÜbern
        End Get
        Set(value As String)
            _planeleÜbern = value
        End Set
    End Property
    ''' <summary>
    ''' Liest und schreibt die Hierarchie, wie sie in ProjektTafel zu finden ist
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Property PThierarchie As String
        Get
            PThierarchie = _PThierarchie
        End Get
        Set(value As String)
            _PThierarchie = value
        End Set
    End Property
    ''' <summary>
    ''' Liest und schreibt den Grund der Übernahme ins Projekt
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Property grund As String
        Get
            grund = _grund
        End Get
        Set(value As String)
            _grund = value
        End Set
    End Property
    ''' <summary>
    ''' Liest und schreibt Name der Darstellungsklasse in ProjektTafel
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Property PTklasse As String
        Get
            PTklasse = _PTklasse
        End Get
        Set(value As String)
            _PTklasse = value
        End Set
    End Property


    ''' <summary>
    ''' erzeugt ein neues Element der Klasse clsLogbuchline
    ''' </summary>
    ''' <remarks></remarks>
    Sub New()
        _actDate = Date.Now.ToString
        _Projekt = ""
        _hierarchie = ""
        _planelement = ""
        _klasse = ""
        _abkürzung = ""
        _quelle = ""
        _planeleÜbern = ""
        _grund = ""
        _PThierarchie = ""
        _PTklasse = ""
    End Sub


    ''' <summary>
    ''' initialisert im Inputfile die Tabelle 'Logbuch'
    ''' 
    ''' </summary>
    ''' <remarks></remarks>
    Sub InitProtokoll(ByRef wslogbuch As xlNS.Worksheet, ByVal tabblattname As String)

        Try
            wslogbuch = CType(xlsLogfile.Worksheets(tabblattname), _
               Global.Microsoft.Office.Interop.Excel.Worksheet)


            If Not IsNothing(wslogbuch) Then

                xlsLogfile.Worksheets.Application.DisplayAlerts = False
                wslogbuch.Delete()
                xlsLogfile.Worksheets.Application.DisplayAlerts = True

                wslogbuch = CType(xlsLogfile.Worksheets.Add(), _
                   Global.Microsoft.Office.Interop.Excel.Worksheet)
                wslogbuch.Name = tabblattname
            End If
        Catch ex As Exception
            'wsLogbuch = CType(xlsInput.Worksheets.Add(After:=xlsInput.Worksheets.Count), _
            '   Global.Microsoft.Office.Interop.Excel.Worksheet)
            wslogbuch = CType(xlsLogfile.Worksheets.Add(), _
                Global.Microsoft.Office.Interop.Excel.Worksheet)
            wslogbuch.Name = tabblattname
        End Try


        With wslogbuch
            .Rows.RowHeight = 15
            CType(.Rows(1), xlNS.Range).RowHeight = 30
            CType(.Rows(1), xlNS.Range).Font.Bold = True

            If awinSettings.fullProtocol Then
                CType(.Cells(1, 1), xlNS.Range).Value() = "Datum"
                CType(.Cells(1, 2), xlNS.Range).Value() = "Projekt"
                CType(.Cells(1, 3), xlNS.Range).Value() = "Hierarchie"
                CType(.Cells(1, 4), xlNS.Range).Value() = "Plan-Element"
                CType(.Cells(1, 5), xlNS.Range).Value() = "Klasse"
                CType(.Cells(1, 6), xlNS.Range).Value() = "Abkürzung"
                CType(.Cells(1, 7), xlNS.Range).Value() = "Quelle"
                CType(.Cells(1, 8), xlNS.Range).Value() = "Übernommen als"
                CType(.Cells(1, 9), xlNS.Range).Value() = "Grund"
                CType(.Cells(1, 10), xlNS.Range).Value() = "PT Hierarchie"
                CType(.Cells(1, 11), xlNS.Range).Value() = "PT Klasse"
                CType(.Columns(1), xlNS.Range).ColumnWidth = 40
                CType(.Columns(2), xlNS.Range).ColumnWidth = 40
                CType(.Columns(3), xlNS.Range).ColumnWidth = 40
                CType(.Columns(4), xlNS.Range).ColumnWidth = 40
                CType(.Columns(5), xlNS.Range).ColumnWidth = 40
                CType(.Columns(6), xlNS.Range).ColumnWidth = 40
                CType(.Columns(7), xlNS.Range).ColumnWidth = 40
                CType(.Columns(8), xlNS.Range).ColumnWidth = 40
                CType(.Columns(9), xlNS.Range).ColumnWidth = 40
                CType(.Columns(10), xlNS.Range).ColumnWidth = 40
                CType(.Columns(11), xlNS.Range).ColumnWidth = 40
            Else
                CType(.Cells(1, 1), xlNS.Range).Value() = "Datum"
                CType(.Cells(1, 8), xlNS.Range).Value() = "Übernommen als"
                CType(.Cells(1, 9), xlNS.Range).Value() = "Grund"
                CType(.Columns(1), xlNS.Range).ColumnWidth = 40
                CType(.Columns(8), xlNS.Range).ColumnWidth = 40
                CType(.Columns(9), xlNS.Range).ColumnWidth = 40

            End If

        End With

        Me.tabblattname = tabblattname
    End Sub

    ''' <summary>
    ''' Schreibt eine Zeile in die Tabelle 'Logbuch' der Input-Datei
    ''' </summary>
    ''' <remarks></remarks>
    Sub writeLog(ByRef rowoffset As Integer)


        Dim zelle As xlNS.Range

        Dim wsLogbuch As xlNS.Worksheet = Nothing

        Try
            wsLogbuch = CType(xlsLogfile.Worksheets(Me.tabblattname), _
               Global.Microsoft.Office.Interop.Excel.Worksheet)


        Catch ex As Exception

            InitProtokoll(wsLogbuch, Me.tabblattname) ' Tabelle Logbuch wird initialisiert
            rowoffset = 3
            If Not IsNothing(xlsLogfile) Then
                xlsLogfile.Save()
            End If
        End Try

        wsLogbuch.Unprotect(Password:="x")       ' Schreibschutz  für Logbuch aufheben

        Try
            'rowOffset = CType(CType(xlsLogfile.Worksheets(Me.tabblattname), xlNS.Worksheet).Cells(20000, 1), Global.Microsoft.Office.Interop.Excel.Range).End(XlDirection.xlUp).Row
            rowoffset = rowoffset + 1
            zelle = CType(wsLogbuch.Rows(rowoffset), xlNS.Range)
            With zelle
                If awinSettings.fullProtocol Then

                    CType(.Cells(1, 1), xlNS.Range).Value = _actDate
                    CType(.Cells(1, 2), xlNS.Range).Value = _Projekt
                    CType(.Cells(1, 3), xlNS.Range).Value = _hierarchie
                    CType(.Cells(1, 4), xlNS.Range).Value = _planelement
                    CType(.Cells(1, 5), xlNS.Range).Value = _klasse
                    CType(.Cells(1, 6), xlNS.Range).Value = _abkürzung
                    CType(.Cells(1, 7), xlNS.Range).Value = _quelle
                    CType(.Cells(1, 8), xlNS.Range).Value = _planeleÜbern
                    CType(.Cells(1, 9), xlNS.Range).Value = _grund
                    CType(.Cells(1, 10), xlNS.Range).Value = _PThierarchie
                    CType(.Cells(1, 11), xlNS.Range).Value = _PTklasse
                Else
                    CType(.Cells(1, 1), xlNS.Range).Value = _actDate
                    CType(.Cells(1, 8), xlNS.Range).Value = _planeleÜbern
                    CType(.Cells(1, 9), xlNS.Range).Value = _grund
                End If
            End With
        Catch ex As Exception

        End Try

        ' Schreibschutz wieder setzen
        wsLogbuch.Protect(Password:="x", UserInterfaceOnly:=True, DrawingObjects:=True, Contents:=True, Scenarios:=True)

        '' '' Logbuch sichern
        ' ''If Not IsNothing(xlsLogfile) Then
        ' ''    xlsLogfile.Save()
        ' ''End If


    End Sub
    Sub close()
        ' Logbuch sichern
        If Not IsNothing(xlsLogfile) Then
            xlsLogfile.Save()
        End If

    End Sub
    Sub clear()
        _actDate = ""
        _Projekt = ""
        _hierarchie = ""
        _planelement = ""
        _klasse = ""
        _abkürzung = ""
        _quelle = ""
        _planeleÜbern = ""
        _grund = ""
        _PThierarchie = ""
        _PTklasse = ""
    End Sub

End Class

