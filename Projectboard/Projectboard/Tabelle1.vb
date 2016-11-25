Imports System.Math ' für Funktion Abs()
Imports ProjectBoardDefinitions
Imports ProjectBoardBasic
Imports Microsoft.Office.Interop.Excel
Imports Microsoft.Office.Core



Public Class Tabelle1

    Dim projektfarbe As Object, projektschrift As Integer

    Private Sub Tabelle1_ActivateEvent() Handles Me.ActivateEvent
        Dim a As Integer


        Try
            Application.DisplayFormulaBar = False

            With Application.ActiveWindow

                .DisplayWorkbookTabs = False
                .DisplayHeadings = False

                If .SplitRow = 1 Then
                    ' nichts tun
                Else
                    .SplitColumn = 0
                    .SplitRow = 1
                End If
              

                .GridlineColor = RGB(220, 220, 220)
                a = Application.ActiveWindow.Panes.Count

                .FreezePanes = True

            End With


        Catch ex As Exception
            ' nur eine Dummy Zuweisung, um ggf später hier einen Haltepunkt setzen zu können
            Dim b As Integer = a
        End Try

        If appInstance.ScreenUpdating = False Then
            appInstance.ScreenUpdating = True
        End If

    End Sub

    Private Sub Tabelle1_Startup() Handles Me.Startup

        Application.DisplayFormulaBar = False

        With Application.ActiveWindow
            .SplitColumn = 0
            .SplitRow = 1
            .DisplayWorkbookTabs = False
            .GridlineColor = RGB(220, 220, 220)
            .FreezePanes = True
            .DisplayHeadings = False
            .Zoom = 100
        End With
    End Sub


  

    Private Sub Tabelle1_BeforeRightClick(Target As Microsoft.Office.Interop.Excel.Range, ByRef Cancel As Boolean) Handles Me.BeforeRightClick

        ' damit wird der Standard Right Click deaktiviert
        Cancel = True

    End Sub

    Private Sub Tabelle1_Change(Target As Microsoft.Office.Interop.Excel.Range) Handles Me.Change

        appInstance.EnableEvents = False

        Try
            Target.Clear()
        Catch ex As Exception

        End Try
        appInstance.EnableEvents = True

    End Sub

    Private Sub Tabelle1_Deactivate() Handles Me.Deactivate




    End Sub

    Private Sub Tabelle1_SelectionChange(Target As Microsoft.Office.Interop.Excel.Range) Handles Me.SelectionChange
        Dim stdLaenge As Integer = 12

        Dim pname As String = ""

        Dim selectTop As Double, selectLeft As Double, selectWidth As Double, selectHeight As Double

        Dim hproj As New clsProjekt
        Dim tmpShapes As Excel.Shapes
        Dim shpArray() As String
        Dim tmpArray() As String
        Dim anzahlSel As Integer = 0

        Dim von As Integer, bis As Integer

        Dim c1 As Range, c2 As Range
        Dim maxRows As Integer = CType(Application.ActiveSheet, Excel.Worksheet).Rows.Count
        Dim maxColumns As Integer = CType(Application.ActiveSheet, Excel.Worksheet).Columns.Count


        Dim eingabebereich As Excel.Range
        Dim kalenderbereich As Excel.Range

        With CType(Application.ActiveSheet, Excel.Worksheet)
            eingabebereich = .Range(.Cells(2, 1), .Cells(maxRows, maxColumns))
            kalenderbereich = .Range(.Cells(1, 1), .Cells(1, maxColumns))
        End With

        c1 = Application.Intersect(Target, eingabebereich)
        c2 = Application.Intersect(Target, kalenderbereich)

        appInstance.EnableEvents = False


        ' Die Selektierten Projekte zurücksetzen 

        If selectedProjekte.Count > 0 Then
            selectedProjekte.Clear()
            Call awinNeuZeichnenDiagramme(8)
        End If

        tmpShapes = CType(CType(appInstance.ActiveSheet, Excel.Worksheet).Shapes, Excel.Shapes)



        If Not c2 Is Nothing Then
            c1 = Nothing

            von = c2.Column
            bis = von + c2.Columns.Count - 1
            Call awinChangeTimeSpan(von, bis)
            Call awinDeSelect()

            'If c2.Columns.Count > 5 Then
            '    von = c2.Column
            '    bis = von + c2.Columns.Count - 1
            '    Call awinChangeTimeSpan(von, bis)
            '    Call awinDeSelect()
            'Else
            '    Call awinDeSelect()
            'End If

        Else

            If c1 Is Nothing Then

            ElseIf Target.Rows.Count > 1 Or Target.Columns.Count > 1 Then
                ' multi-Select 

                Dim formerSU As Boolean = appInstance.ScreenUpdating
                appInstance.ScreenUpdating = False

                With Target
                    selectTop = .Row - 1
                    selectLeft = .Column
                    selectWidth = .Columns.Count
                    selectHeight = .Rows.Count
                End With


                ReDim tmpArray(ShowProjekte.Count - 1)

                For Each kvp As KeyValuePair(Of String, clsProjekt) In ShowProjekte.Liste

                    With kvp.Value
                        If (.tfspalte >= selectLeft) And (.tfspalte + .anzahlRasterElemente <= selectLeft + selectWidth) _
                            And (.tfZeile >= selectTop) And (.tfZeile <= selectTop + selectHeight - 1) Then

                            ' ist in der Range - also selektieren
                            anzahlSel = anzahlSel + 1
                            tmpArray(anzahlSel - 1) = kvp.Value.name
                        End If
                    End With

                Next
                If anzahlSel > 0 Then
                    Try
                        ReDim shpArray(anzahlSel - 1)
                        shpArray = tmpArray
                        tmpShapes.Range(shpArray).Select()
                    Catch ex As Exception

                    End Try

                End If

                appInstance.ScreenUpdating = formerSU
            Else
                ' nichts tun 
            End If
        End If


        appInstance.EnableEvents = True


    End Sub

    
End Class
