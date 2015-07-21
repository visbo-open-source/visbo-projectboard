Imports Microsoft.Office.Interop
Imports Microsoft.Office.Interop.Excel
Imports System.Windows.Forms
Imports Microsoft.Office.Core
Imports xlNS = Microsoft.Office.Interop.Excel
Imports System.ComponentModel
Imports Microsoft.VisualBasic.Constants


Public Module Projekte


    ''' <summary>
    ''' zeigt den Vergleich zwischen zwei Items an; Ausgang können zwei Projekt-Varianten sein, aber 
    ''' auch der Vergleich zwischen zwei Konstellationen zum Beispiel 
    ''' </summary>
    ''' <param name="name1">1. Bezeichner des Vergleiches </param>
    ''' <param name="values1">Werte, bezogen auf comparisonItem, für den 1. Bezeichner</param>
    ''' <param name="name2">2. Bezeichner des Vergleiches</param>
    ''' <param name="values2">Werte, bezogen auf comparisonItem, für den 2. Bezeichner</param>
    ''' <param name="comparisonItem">wofür stehen die Werte: Rolle, Kostenart, Kennzahl, etc. </param>
    ''' <param name="massEinheit">meist entweder T€ oder MM</param>
    ''' <param name="top">Platzierungs-Info für das Chart</param>
    ''' <param name="left">Platzierungs-Info für das Chart</param>
    ''' <param name="width">Platzierungs-Info für das Chart</param>
    ''' <param name="height">Platzierungs-Info für das Chart</param>
    ''' <remarks></remarks>
    Public Sub ShowDiagramCompare(ByRef name1 As String, ByRef values1() As Double, ByRef name2 As String, ByRef values2() As Double, ByRef comparisonItem As String, ByRef massEinheit As String, _
                               ByVal top As Double, ByVal left As Double, ByVal width As Double, ByVal height As Double)
        Dim diagramtitle As String, chtTitle As String
        Dim sum1 As Double, sum2 As Double, diff As Double
        Dim array1() As Double, array2() As Double
        Dim maxlength As Integer = System.Math.Max(values1.Length, values2.Length)
        Dim minlength As Integer = System.Math.Min(values1.Length, values2.Length)
        Dim maxscale As Integer
        Dim i As Integer
        Dim Xdatenreihe() As String
        Dim anzDiagrams As Integer


        Dim majorUnit As Integer

        Dim maxValue As Double = System.Math.Max(values1.Max, values2.Max)
        height = System.Math.Max(40 + System.Math.Log(maxValue), 100)
        'majorUnit = System.Math.Max(maxValue / 5, 2)
        maxscale = CInt(System.Math.Max(10, maxValue * 1.3))

        If maxscale < 100 Then
            maxscale = CInt(System.Math.Round(maxscale / 10, MidpointRounding.ToEven) * 10)
        Else
            maxscale = CInt(System.Math.Round(maxscale / 100, MidpointRounding.ToEven) * 100)
        End If

        If maxscale < 10 Then maxscale = 10
        majorUnit = CInt(maxscale / 4)


        

        Dim formerEE As Boolean = appInstance.EnableEvents
        Dim formerSU As Boolean = appInstance.ScreenUpdating


        ReDim Xdatenreihe(maxlength - 1)

        ReDim array1(maxlength - 1)
        ReDim array2(maxlength - 1)

        ' Format(diff, "##,##0")
        sum1 = values1.Sum
        sum2 = values2.Sum
        diff = sum1 - sum2

        For i = 1 To maxlength
            'Xdatenreihe(i - 1) = StartofCalendar.AddMonths(pstart + i - 2).ToString("MMM yy")
            Xdatenreihe(i - 1) = i.ToString
        Next i


        If name2 = "" Then ' es handelt sich um den Vergleich mit der Beauftragung
            
            diagramtitle = comparisonItem & " ( " & Format(sum1, "#,###,0") & " - " & Format(sum2, "#,###,0") & " = " & Format(diff, "###,0") & " " & massEinheit & " )"

        Else

            diagramtitle = comparisonItem & " ( " & Format(sum1, "#,###,0") & " - " & Format(sum2, "#,###,0") & " = " & Format(diff, "###,0") & " " & massEinheit & " )" 

        End If

        ' Kopieren der Werte in Array1
        For i = 0 To values1.Length - 1
            array1(i) = values1(i)
        Next i

        For i = 0 To values2.Length - 1
            array2(i) = values2(i)
        Next i


        Dim found As Boolean

        If formerEE Then
            appInstance.EnableEvents = False
        End If

        If formerSU Then
            appInstance.ScreenUpdating = False
        End If

        With CType(appInstance.Worksheets(arrWsNames(3)), Excel.Worksheet)
            anzDiagrams = CType(.ChartObjects, Excel.ChartObjects).Count
            '
            ' um welches Diagramm handelt es sich ...
            '
            i = 1
            found = False
            While i <= anzDiagrams And Not found
                Try
                    chtTitle = CType(.ChartObjects(i), Excel.ChartObject).Chart.ChartTitle.Text
                Catch ex As Exception
                    chtTitle = " "
                End Try
                If chtTitle = diagramtitle Then
                    found = True
                Else
                    i = i + 1
                End If

            End While

            If found Then
                appInstance.EnableEvents = formerEE
                appInstance.ScreenUpdating = formerSU
                Throw New ArgumentException("Diagramm wird schon angezeigt ...")
            Else

                With CType(appInstance.Charts.Add, Excel.Chart)
                    ' remove extra series
                    Do Until CType(.SeriesCollection, Excel.SeriesCollection).Count = 0
                        CType(.SeriesCollection(1), Excel.Series).Delete()
                    Loop



                    With CType(.SeriesCollection, Excel.SeriesCollection).NewSeries
                        .Name = name1
                        .Interior.Color = vergleichsfarbe1
                        .Values = array1
                        .XValues = Xdatenreihe
                        ' Unterschied farblich hervorheben ...
                        For ix = 1 To maxlength
                            If array1(ix - 1) = array2(ix - 1) Then
                                With CType(.Points(ix), Excel.Point)
                                    .Interior.Color = vergleichsfarbe0
                                End With
                            End If
                        Next
                        .ChartType = Excel.XlChartType.xlColumnClustered
                    End With

                    With CType(.SeriesCollection, Excel.SeriesCollection).NewSeries
                        .Name = name2
                        .Interior.Color = vergleichsfarbe2
                        .Values = array2
                        .XValues = Xdatenreihe
                        ' Unterschied farblich hervorheben ...
                        For ix = 1 To maxlength
                            If array1(ix - 1) = array2(ix - 1) Then
                                With CType(.Points(ix), Excel.Point)
                                    .Interior.Color = vergleichsfarbe0
                                End With
                            End If
                        Next
                        .ChartType = Excel.XlChartType.xlColumnClustered
                    End With


                    .HasAxis(Excel.XlAxisType.xlCategory) = True
                    .HasAxis(Excel.XlAxisType.xlValue) = True

                    'für test-zwecke 
                    'Dim ax As Excel.Axes
                    'With ax
                    '    .Item(Excel.XlAxisType.xlCategory).Format.TextFrame2.TextRange.Font.Size = 8
                    '    .Item(Excel.XlAxisType.xlValue).
                    'End With

                    With CType(.Axes(Excel.XlAxisType.xlCategory), Excel.Axis)
                        .HasTitle = False
                        '.MinimumScale = 0

                        '.Format.TextFrame2.TextRange.Font.Size = MsoAutoSize.msoAutoSizeTextToFitShape
                        'With .AxisTitle
                        '    .Characters.text = "Monate"
                        '    .Font.Size = 8
                        '    '.Font.Size = MsoAutoSize.msoAutoSizeTextToFitShape
                        'End With
                        'Try
                        '    .Format.TextFrame2.TextRange.Font.Size = 8
                        'Catch ex As Exception

                        'End Try
                    End With

                    With CType(.Axes(Excel.XlAxisType.xlValue), Excel.Axis)
                        .HasTitle = False
                        .MinimumScale = 0
                        .HasMinorGridlines = False
                        .HasMajorGridlines = True
                        .MajorUnit = majorUnit

                        '.Format.TextFrame2.TextRange.Font.Size = MsoAutoSize.msoAutoSizeTextToFitShape
                        'With .AxisTitle
                        '    .Characters.text = comparisonItem
                        '    .Font.Size = 8
                        '    '.Font.Size = MsoAutoSize.msoAutoSizeTextToFitShape
                        'End With
                        'Try
                        '    .Format.TextFrame2.TextRange.Font.Size = 8
                        'Catch ex As Exception

                        'End Try

                    End With

                    .HasLegend = True
                    With .Legend
                        .Position = Excel.XlLegendPosition.xlLegendPositionTop
                        .Font.Size = 10
                        '.Font.Size = MsoAutoSize.msoAutoSizeTextToFitShape
                    End With

                    .HasTitle = True
                    With .ChartTitle
                        .Text = diagramtitle
                        .Font.Size = 12
                        '.Font.Size = MsoAutoSize.msoAutoSizeTextToFitShape
                    End With

                    .Location(Where:=XlChartLocation.xlLocationAsObject, Name:=CType(appInstance.Worksheets(arrWsNames(3)), Excel.Worksheet).Name)
                End With

                ' jetzt kommt die Korrektur der Größe; herausfinden, wieviel Raum die Axis Beschriftung einnimmt ... 
                With CType(.ChartObjects(anzDiagrams + 1), Excel.ChartObject)
                    .Top = top
                    .Height = 2 * height

                    Dim axleft As Double, axwidth As Double
                    If CBool(.Chart.HasAxis(Excel.XlAxisType.xlValue)) Then
                        With CType(.Chart.Axes(Excel.XlAxisType.xlValue), Excel.Axis)
                            axleft = .Left
                            axwidth = .Width
                        End With
                        If left - axwidth < 1 Then
                            left = 1
                            width = width + left + 9
                        Else
                            left = left - axwidth
                            width = width + axwidth + 9
                        End If

                    End If

                    .Left = left
                    .Width = width


                End With

                'With .ChartObjects(anzDiagrams + 1)
                '    .top = top
                '    .left = left
                '    .height = height
                '    .width = width
                'End With


            End If


        End With


        appInstance.EnableEvents = formerEE
        appInstance.ScreenUpdating = formerSU

    End Sub

    Public Sub ShowDiagramCompare1(ByRef name1 As String, ByRef values1() As Double, ByRef name2 As String, ByRef values2() As Double, ByRef comparisonItem As String, ByRef massEinheit As String, _
                               ByVal top As Double, ByVal left As Double, ByVal width As Double, ByVal height As Double)
        Dim diagramtitle As String, chtTitle As String
        Dim sum1 As Double, sum2 As Double, diff As Double
        Dim array0() As Double, array1() As Double, array2() As Double
        Dim maxlength As Integer = System.Math.Max(values1.Length, values2.Length)
        Dim minlength As Integer = System.Math.Min(values1.Length, values2.Length)
        Dim maxscale As Integer
        Dim i As Integer
        Dim Xdatenreihe() As String
        Dim anzDiagrams As Integer


        Dim majorUnit As Integer

        Dim maxValue As Double = System.Math.Max(values1.Max, values2.Max)
        height = System.Math.Max(40 + System.Math.Log(maxValue), 100)
        'majorUnit = System.Math.Max(maxValue / 5, 2)
        maxscale = CInt(System.Math.Max(10, maxValue * 1.3))

        If maxscale < 100 Then
            maxscale = CInt(System.Math.Round(maxscale / 10, MidpointRounding.ToEven) * 10)
        Else
            maxscale = CInt(System.Math.Round(maxscale / 100, MidpointRounding.ToEven) * 100)
        End If

        If maxscale < 10 Then maxscale = 10
        majorUnit = CInt(maxscale / 4)




        Dim formerEE As Boolean = appInstance.EnableEvents
        Dim formerSU As Boolean = appInstance.ScreenUpdating


        ReDim Xdatenreihe(maxlength - 1)
        ReDim array0(maxlength - 1)
        ReDim array1(maxlength - 1)
        ReDim array2(maxlength - 1)

        ' Format(diff, "##,##0")
        sum1 = values1.Sum
        sum2 = values2.Sum
        diff = sum1 - sum2

        For i = 1 To maxlength
            'Xdatenreihe(i - 1) = StartofCalendar.AddMonths(pstart + i - 2).ToString("MMM yy")
            Xdatenreihe(i - 1) = i.ToString
        Next i


        If name2 = "" Then ' es handelt sich um den Vergleich mit der Beauftragung
            diagramtitle = comparisonItem & " ( " & Format(sum1, "#,###,0") & " - " & Format(sum2, "#,###,0") & " = " & Format(diff, "###,0") & " " & massEinheit & " )" & vbLf & _
                                                name1 & " versus Beauftragung"
        Else

            diagramtitle = comparisonItem & " ( " & Format(sum1, "#,###,0") & " - " & Format(sum2, "#,###,0") & " = " & Format(diff, "###,0") & " " & massEinheit & " )" & vbLf & _
                                                name1 & " versus " & name2

        End If


        For i = 0 To minlength - 1
            If values1(i) >= values2(i) Then
                array0(i) = values2(i)
                array1(i) = values1(i) - values2(i)
                array2(i) = 0
            Else
                array0(i) = values1(i)
                array1(i) = 0
                array2(i) = values2(i) - values1(i)
            End If
        Next i

        If values1.Length > values2.Length Then
            For i = minlength To maxlength - 1
                array1(i) = values1(i)
            Next
        ElseIf values1.Length < values2.Length Then
            For i = minlength To maxlength - 1
                array2(i) = values2(i)
            Next
        End If

        Dim found As Boolean

        If formerEE Then
            appInstance.EnableEvents = False
        End If

        If formerSU Then
            appInstance.ScreenUpdating = False
        End If

        With CType(appInstance.Worksheets(arrWsNames(3)), Excel.Worksheet)
            anzDiagrams = CType(.ChartObjects, Excel.ChartObjects).Count
            '
            ' um welches Diagramm handelt es sich ...
            '
            i = 1
            found = False
            While i <= anzDiagrams And Not found
                Try
                    chtTitle = CType(.ChartObjects(i), Excel.ChartObject).Chart.ChartTitle.Text
                Catch ex As Exception
                    chtTitle = " "
                End Try
                If chtTitle = diagramtitle Then
                    found = True
                Else
                    i = i + 1
                End If

            End While

            If found Then
                appInstance.EnableEvents = formerEE
                appInstance.ScreenUpdating = formerSU
                Throw New ArgumentException("Diagramm wird schon angezeigt ...")
            Else

                With CType(appInstance.Charts.Add, Excel.Chart)
                    ' remove extra series
                    Do Until CType(.SeriesCollection, Excel.SeriesCollection).Count = 0
                        CType(.SeriesCollection(1), Excel.Series).Delete()
                    Loop


                    'series
                    With CType(.SeriesCollection, Excel.SeriesCollection).NewSeries
                        .Name = "identisch"
                        .Interior.Color = vergleichsfarbe0
                        .Values = array0
                        .XValues = Xdatenreihe
                        .ChartType = Excel.XlChartType.xlColumnStacked
                    End With


                    With CType(.SeriesCollection, Excel.SeriesCollection).NewSeries
                        '.name = "mehr"
                        .Name = name1
                        .Interior.Color = vergleichsfarbe1
                        .Values = array1
                        .XValues = Xdatenreihe
                        .ChartType = Excel.XlChartType.xlColumnStacked
                    End With

                    With CType(.SeriesCollection, Excel.SeriesCollection).NewSeries
                        '.name = "weniger"
                        .Name = name2
                        .Interior.Color = vergleichsfarbe2
                        .Values = array2
                        .XValues = Xdatenreihe
                        .ChartType = Excel.XlChartType.xlColumnStacked
                    End With


                    .HasAxis(Excel.XlAxisType.xlCategory) = True
                    .HasAxis(Excel.XlAxisType.xlValue) = True

                    'für test-zwecke 
                    'Dim ax As Excel.Axes
                    'With ax
                    '    .Item(Excel.XlAxisType.xlCategory).Format.TextFrame2.TextRange.Font.Size = 8
                    '    .Item(Excel.XlAxisType.xlValue).
                    'End With

                    With CType(.Axes(Excel.XlAxisType.xlCategory), Excel.Axis)
                        .HasTitle = False
                        '.MinimumScale = 0

                        '.Format.TextFrame2.TextRange.Font.Size = MsoAutoSize.msoAutoSizeTextToFitShape
                        'With .AxisTitle
                        '    .Characters.text = "Monate"
                        '    .Font.Size = 8
                        '    '.Font.Size = MsoAutoSize.msoAutoSizeTextToFitShape
                        'End With
                        'Try
                        '    .Format.TextFrame2.TextRange.Font.Size = 8
                        'Catch ex As Exception

                        'End Try
                    End With

                    With CType(.Axes(Excel.XlAxisType.xlValue), Excel.Axis)
                        .HasTitle = False
                        .MinimumScale = 0
                        .HasMinorGridlines = False
                        .HasMajorGridlines = True
                        .MajorUnit = majorUnit

                        '.Format.TextFrame2.TextRange.Font.Size = MsoAutoSize.msoAutoSizeTextToFitShape
                        'With .AxisTitle
                        '    .Characters.text = comparisonItem
                        '    .Font.Size = 8
                        '    '.Font.Size = MsoAutoSize.msoAutoSizeTextToFitShape
                        'End With
                        'Try
                        '    .Format.TextFrame2.TextRange.Font.Size = 8
                        'Catch ex As Exception

                        'End Try

                    End With

                    .HasLegend = True
                    With .Legend
                        .Position = Excel.XlLegendPosition.xlLegendPositionTop
                        .Font.Size = 10
                        '.Font.Size = MsoAutoSize.msoAutoSizeTextToFitShape
                    End With

                    .HasTitle = True
                    With .ChartTitle
                        .Text = diagramtitle
                        .Font.Size = 12
                        '.Font.Size = MsoAutoSize.msoAutoSizeTextToFitShape
                    End With

                    .Location(Where:=XlChartLocation.xlLocationAsObject, Name:=CType(appInstance.Worksheets(arrWsNames(3)), Excel.Worksheet).Name)
                End With

                ' jetzt kommt die Korrektur der Größe; herausfinden, wieviel Raum die Axis Beschriftung einnimmt ... 
                With CType(.ChartObjects(anzDiagrams + 1), Excel.ChartObject)
                    .Top = top
                    .Height = 2 * height

                    Dim axleft As Double, axwidth As Double
                    If CBool(.Chart.HasAxis(Excel.XlAxisType.xlValue)) Then
                        With CType(.Chart.Axes(Excel.XlAxisType.xlValue), Excel.Axis)
                            axleft = .Left
                            axwidth = .Width
                        End With
                        If left - axwidth < 1 Then
                            left = 1
                            width = width + left + 9
                        Else
                            left = left - axwidth
                            width = width + axwidth + 9
                        End If

                    End If

                    .Left = left
                    .Width = width


                End With

                'With .ChartObjects(anzDiagrams + 1)
                '    .top = top
                '    .left = left
                '    .height = height
                '    .width = width
                'End With


            End If


        End With


        appInstance.EnableEvents = formerEE
        appInstance.ScreenUpdating = formerSU

    End Sub


    ''' <summary>
    ''' zeigt den Vergleich zwischen zwei Items an; Ausgang können zwei Projekt-Varianten sein, aber 
    ''' auch der Vergleich zwischen zwei Konstellationen zum Beispiel ; 
    ''' gezeigt wird nur das Delta , nicht auch der identische Anteil 
    ''' </summary>
    ''' <param name="name1">1. Bezeichner des Vergleiches </param>
    ''' <param name="values1">Werte, bezogen auf comparisonItem, für den 1. Bezeichner</param>
    ''' <param name="name2">2. Bezeichner des Vergleiches</param>
    ''' <param name="values2">Werte, bezogen auf comparisonItem, für den 2. Bezeichner</param>
    ''' <param name="comparisonItem">wofür stehen die Werte: Rolle, Kostenart, Kennzahl, etc. </param>
    ''' <param name="massEinheit">meist entweder T€ oder MM</param>
    ''' <param name="top">Platzierungs-Info für das Chart</param>
    ''' <param name="left">Platzierungs-Info für das Chart</param>
    ''' <param name="width">Platzierungs-Info für das Chart</param>
    ''' <param name="height">Platzierungs-Info für das Chart</param>
    ''' <remarks></remarks>
    Public Sub ShowDiagramCompare2(ByRef name1 As String, ByRef values1() As Double, ByRef name2 As String, ByRef values2() As Double, ByRef comparisonItem As String, ByRef massEinheit As String, _
                               ByVal top As Double, ByVal left As Double, ByVal width As Double, ByVal height As Double)
        ' wenn das weider verwendet wird: erst anpassen Berechnung left und width  wie in compareDiagram

        Dim diagramtitle As String, chtTitle As String
        Dim sum1 As Double
        Dim array0() As Double, array1() As Double, array2() As Double
        Dim maxlength As Integer = System.Math.Max(values1.Length, values2.Length)
        Dim minlength As Integer = System.Math.Min(values1.Length, values2.Length)
        Dim i As Integer
        Dim Xdatenreihe() As String
        Dim anzDiagrams As Integer

        Dim formerEE As Boolean = appInstance.EnableEvents
        Dim formerSU As Boolean = appInstance.ScreenUpdating



        'Try
        '    pstart = ShowProjekte.getProject(name1).Start
        'Catch ex As Exception
        '    Call MsgBox("Fehler in ShowDiagramCompare beim Auslesen des Starts...")
        '    Exit Sub
        'End Try

        ReDim Xdatenreihe(maxlength - 1)
        ReDim array0(maxlength - 1)
        ReDim array1(maxlength - 1)
        ReDim array2(maxlength - 1)


        'sum1 = values1.Sum
        'sum2 = values2.Sum

        'For i = 1 To maxlength
        '    Xdatenreihe(i - 1) = "" & i.ToString
        'Next

        For i = 1 To maxlength
            'Xdatenreihe(i - 1) = StartofCalendar.AddMonths(pstart + i - 2).ToString("MMM yy")
            Xdatenreihe(i - 1) = i.ToString
        Next i

        For i = 0 To minlength - 1
            array1(i) = values1(i) - values2(i)
        Next i

        If values1.Length > values2.Length Then
            For i = minlength To maxlength - 1
                array1(i) = values1(i)
            Next
        ElseIf values1.Length < values2.Length Then
            For i = minlength To maxlength - 1
                array1(i) = values2(i) * -1
            Next
        End If

        sum1 = array1.Sum
        diagramtitle = "Übersicht Mehr- bzw. Minder-Aufwände im Vergleich zur Beauftragung/Vergleichsprojekt" & vbLf _
                       & name1 & " (" & sum1.ToString & " " & massEinheit & ") " & vbLf _
                       & "bezogen auf " & comparisonItem

        ' Format(maxwert, "##,##0")

        Dim found As Boolean

        If formerEE Then
            appInstance.EnableEvents = False
        End If

        If formerSU Then
            appInstance.ScreenUpdating = False
        End If

        With CType(appInstance.Worksheets(arrWsNames(3)), Excel.Worksheet)
            anzDiagrams = CType(.ChartObjects, Excel.ChartObjects).Count
            '
            ' um welches Diagramm handelt es sich ...
            '
            i = 1
            found = False
            While i <= anzDiagrams And Not found
                Try
                    chtTitle = CType(.ChartObjects(i), Excel.ChartObject).Chart.ChartTitle.Text
                Catch ex As Exception
                    chtTitle = " "
                End Try
                If chtTitle = diagramtitle Then
                    found = True
                Else
                    i = i + 1
                End If

            End While

            If found Then
                MsgBox(" Diagramm wird bereits angezeigt ...")
            Else

                With CType(appInstance.Charts.Add, Excel.Chart)
                    ' remove extra series
                    Do Until CType(.SeriesCollection, Excel.SeriesCollection).Count = 0
                        CType(.SeriesCollection(1), Excel.Series).Delete()
                    Loop


                    'series
                    'With .SeriesCollection.NewSeries
                    '    .name = "identischer Teil"
                    '    .Interior.color = vergleichsfarbe0
                    '    .Values = array0
                    '    .XValues = Xdatenreihe
                    '    .ChartType = Excel.XlChartType.xlColumnStacked
                    'End With

                    With CType(.SeriesCollection, Excel.SeriesCollection).NewSeries
                        .Name = name1
                        .Interior.Color = vergleichsfarbe1
                        .Values = array1
                        .XValues = Xdatenreihe
                        .ChartType = Excel.XlChartType.xlColumnClustered
                    End With

                    'With .SeriesCollection.NewSeries
                    '    .name = name2 & " ist größer"
                    '    .Interior.color = vergleichsfarbe2
                    '    .Values = array2
                    '    .XValues = Xdatenreihe
                    '    .ChartType = Excel.XlChartType.xlColumnStacked
                    'End With


                    .HasAxis(Excel.XlAxisType.xlCategory) = False
                    .HasAxis(Excel.XlAxisType.xlValue) = True

                    With CType(.Axes(Excel.XlAxisType.xlCategory), Excel.Axis)
                        .HasTitle = True
                        '.MinimumScale = 0
                        With .AxisTitle
                            .Characters.Text = "Monate"
                            .Font.Size = 8
                        End With
                    End With

                    With CType(.Axes(Excel.XlAxisType.xlValue), Excel.Axis)
                        .HasTitle = True
                        '.MinimumScale = 0
                        With .AxisTitle
                            .Characters.Text = comparisonItem
                            .Font.Size = 8
                        End With
                    End With

                    .HasLegend = True
                    With .Legend
                        .Position = Excel.XlLegendPosition.xlLegendPositionTop
                        .Font.Size = 8
                    End With

                    .HasTitle = True
                    With .ChartTitle
                        .Text = diagramtitle
                        .Font.Size = 10
                    End With

                    .Location(Where:=XlChartLocation.xlLocationAsObject, Name:=CType(appInstance.Worksheets(arrWsNames(3)), Excel.Worksheet).Name)
                End With

                With CType(.ChartObjects(anzDiagrams + 1), Excel.ChartObject)
                    .Top = top
                    .Left = left
                    .Height = height
                    .Width = width
                End With


            End If


        End With


        appInstance.EnableEvents = formerEE
        appInstance.ScreenUpdating = formerSU

    End Sub

    ''' <summary>
    ''' zeigt die Phasen Verläufe eines Projektes an
    ''' </summary>
    ''' <param name="noColorCollection" >enthält die Namen der Phasen, die ohne Farbe dargestellt werden sollen</param>
    ''' <param name="hproj"></param>
    ''' <param name="repObj"></param>
    ''' <param name="maxscale"></param>
    ''' <param name="top"></param>
    ''' <param name="left"></param>
    ''' <param name="height"></param>
    ''' <param name="width"></param>
    ''' <param name="auswahl"></param>
    ''' <remarks>wenn hier etwas geändert wird, muss auch in updatePhasesBalken geändert werden ... 
    ''' </remarks>
    Public Sub createPhasesBalken(ByVal noColorCollection As Collection, ByVal hproj As clsProjekt, ByRef repObj As Excel.ChartObject, ByVal maxscale As Double, _
                                      ByVal top As Double, ByVal left As Double, ByVal height As Double, ByVal width As Double, ByVal auswahl As Integer)
        Dim diagramTitle As String
        Dim anzDiagrams As Integer
        Dim anzPhasen As Integer
        Dim found As Boolean
        Dim plenInDays As Integer
        Dim i As Integer
        Dim Xdatenreihe() As String
        Dim valueColor() As Object
        Dim tdatenreihe1() As Double, mdatenreihe() As Double, tdatenreihe2() As Double, tdatenreihe3() As Double
        Dim kennung As String
        Dim titelTeile(1) As String
        Dim titelTeilLaengen(1) As Integer
        Dim noFarbe As Long = awinSettings.AmpelNichtBewertet
        Dim tmpcollection As New Collection
        Dim pName As String = hproj.name



        Dim formerEE As Boolean = appInstance.EnableEvents
        'Dim formerSU As Boolean = appInstance.ScreenUpdating
        appInstance.EnableEvents = False
        'appInstance.ScreenUpdating = False

        tmpcollection.Add(hproj.getShapeText & "#" & auswahl.ToString)
        kennung = calcChartKennung("pr", PTprdk.Phasen, tmpcollection)

        Try
            If auswahl = PThis.vorlage Then
                titelTeile(0) = "Vorlage " & hproj.VorlagenName & vbLf
                titelTeile(1) = " "
                'kennung = hproj.VorlagenName.Trim & "#Phasen#1"



            ElseIf auswahl = PThis.beauftragung Then
                titelTeile(0) = "Beauftragung " & hproj.startDate.ToShortDateString & _
                                     " - " & hproj.startDate.AddDays(hproj.dauerInDays - 1).ToShortDateString & vbLf
                titelTeile(1) = " (" & hproj.timeStamp.ToString & ") "
                'kennung = hproj.name.Trim & "Beauftragung" & "#Phasen#1"

            ElseIf auswahl = PThis.letzterStand Then
                titelTeile(0) = "letzter Stand " & hproj.startDate.ToShortDateString & _
                                     " - " & hproj.startDate.AddDays(hproj.dauerInDays - 1).ToShortDateString & vbLf
                titelTeile(1) = " (" & hproj.timeStamp.ToString & ") "
                'kennung = hproj.name.Trim & "letzter Stand" & "#Phasen#1"

            Else
                titelTeile(0) = hproj.getShapeText & " ,  " & hproj.startDate.ToShortDateString & _
                                     " - " & hproj.startDate.AddDays(hproj.dauerInDays - 1).ToShortDateString & vbLf

                titelTeile(1) = " (" & hproj.timeStamp.ToString & ") "
                'kennung = hproj.name.Trim & "#Phasen#1"
            End If
        Catch ex As Exception
            titelTeile(0) = hproj.getShapeText & vbLf
            titelTeile(1) = " (" & hproj.timeStamp.ToString & ") "
            'kennung = hproj.name.Trim & "#Phasen#1"
        End Try



        titelTeilLaengen(0) = titelTeile(0).Length
        titelTeilLaengen(1) = titelTeile(1).Length

        diagramTitle = titelTeile(0) & titelTeile(1)

        '
        ' hole die Projektdauer
        '
        With hproj
            plenInDays = .dauerInDays
        End With

        '
        ' hole die Anzahl Phasen, die in diesem Projekt vorkommen
        '
        anzPhasen = hproj.Liste.Count

        If anzPhasen <= 1 Then
            'MsgBox("keine Phasen definiert")
            appInstance.EnableEvents = formerEE
            'appInstance.ScreenUpdating = formerSU
            Exit Sub
        End If


        ReDim Xdatenreihe(anzPhasen - 1)
        ReDim tdatenreihe1(anzPhasen - 1)
        ReDim mdatenreihe(anzPhasen - 1)
        ReDim tdatenreihe2(anzPhasen - 1)
        ReDim tdatenreihe3(anzPhasen - 1)


        ReDim valueColor(anzPhasen - 1)


        'ReDim hsum(anzPhasen - 1)


        For i = 1 To anzPhasen

            With hproj.Liste.Item(i - 1)
                Xdatenreihe(i - 1) = .name
                tdatenreihe1(i - 1) = .startOffsetinDays
                mdatenreihe(i - 1) = tdatenreihe1(i - 1) / 365 * 12
                tdatenreihe2(i - 1) = .dauerInDays
                tdatenreihe3(i - 1) = 0

                If noColorCollection.Count > 0 Then
                    ' dann soll untersucht werden, ob die Phase in der noColorCollection ist 
                    If noColorCollection.Contains(.nameID) Then
                        valueColor(i - 1) = awinSettings.AmpelNichtBewertet
                        'valueColor(i - 1) = iProjektFarbe()
                    Else
                        Try
                            valueColor(i - 1) = .Farbe
                        Catch ex As Exception
                            ' dann ist es die Farbe des Projektes
                            valueColor(i - 1) = hproj.farbe
                        End Try

                    End If
                Else
                    Try
                        valueColor(i - 1) = .Farbe
                    Catch ex As Exception
                        ' dann ist es die Farbe des Projektes
                        valueColor(i - 1) = hproj.farbe
                    End Try
                End If


            End With
        Next i



        With CType(appInstance.Worksheets(arrWsNames(3)), Excel.Worksheet)
            anzDiagrams = CType(.ChartObjects, Excel.ChartObjects).Count
            '
            ' um welches Diagramm handelt es sich ...
            '
            i = 1
            found = False
            While i <= anzDiagrams And Not found

                Try
                    If CType(.ChartObjects(i), Excel.ChartObject).Name = kennung Then
                        found = True
                    Else
                        i = i + 1
                    End If
                Catch ex As Exception
                    i = i + 1
                End Try

            End While

            If found Then
                'MsgBox(" Diagramm wird bereits angezeigt ...")
                repObj = CType(.ChartObjects(i), Excel.ChartObject)
            Else

                With CType(appInstance.Charts.Add, Excel.Chart)
                    ' remove extra series
                    Do Until CType(.SeriesCollection, Excel.SeriesCollection).Count = 0
                        CType(.SeriesCollection(1), Excel.Series).Delete()
                    Loop

                    'Aufbau der Series 

                    With CType(.SeriesCollection, Excel.SeriesCollection).NewSeries

                        For i = 0 To anzPhasen - 1
                            mdatenreihe(i) = tdatenreihe1(i) / 365 * 12
                        Next
                        .Name = "null1"
                        .Interior.ColorIndex = -4142
                        .Values = mdatenreihe
                        .XValues = Xdatenreihe
                        .HasDataLabels = False

                        For px = 1 To anzPhasen

                            With CType(.Points(px), Excel.Point)
                                If tdatenreihe1(px - 1) < 90 Then
                                    .HasDataLabel = False
                                Else
                                    .HasDataLabel = True
                                    .DataLabel.Text = hproj.startDate.AddDays(tdatenreihe1(px - 1)).ToShortDateString
                                    .DataLabel.Font.Size = awinSettings.fontsizeItems + 2
                                    If mdatenreihe(px - 1) < 5 Then

                                        Try
                                            .DataLabel.Position = Excel.XlDataLabelPosition.xlLabelPositionOutsideEnd
                                            .DataLabel.Font.Size = awinSettings.fontsizeItems
                                        Catch ex As Exception
                                            .DataLabel.Position = Excel.XlDataLabelPosition.xlLabelPositionInsideEnd
                                        End Try

                                    Else
                                        .DataLabel.Position = Excel.XlDataLabelPosition.xlLabelPositionInsideEnd
                                    End If

                                End If

                            End With

                        Next

                        .ChartType = Excel.XlChartType.xlBarStacked
                    End With

                    With CType(.SeriesCollection, Excel.SeriesCollection).NewSeries

                        For i = 0 To anzPhasen - 1
                            mdatenreihe(i) = tdatenreihe2(i) / 365 * 12
                        Next
                        .Name = "Phasen Zeitraum"
                        .Values = mdatenreihe
                        .XValues = Xdatenreihe

                        .HasDataLabels = True
                        CType(.DataLabels, Excel.DataLabels).Font.Size = awinSettings.fontsizeItems
                        CType(.DataLabels, Excel.DataLabels).Position = Excel.XlDataLabelPosition.xlLabelPositionCenter

                        For i = 1 To anzPhasen
                            With CType(.Points(i), Excel.Point)
                                .Interior.Color = valueColor(i - 1)

                                If mdatenreihe(i - 1) <= 3 Then
                                    .DataLabel.Text = tdatenreihe2(i - 1).ToString
                                Else
                                    .DataLabel.Text = tdatenreihe2(i - 1).ToString & " Tage"
                                End If

                            End With
                        Next


                        .ChartType = Excel.XlChartType.xlBarStacked
                    End With

                    With CType(.SeriesCollection, Excel.SeriesCollection).NewSeries

                        .Name = "null2"
                        .Interior.ColorIndex = -4142
                        .Values = tdatenreihe3
                        .XValues = Xdatenreihe

                        .HasDataLabels = True
                        CType(.DataLabels, Excel.DataLabels).Font.Size = awinSettings.fontsizeItems + 2
                        CType(.DataLabels, Excel.DataLabels).Position = Excel.XlDataLabelPosition.xlLabelPositionInsideBase

                        Dim bis As Integer
                        For px = 1 To anzPhasen

                            With CType(.Points(px), Excel.Point)

                                bis = CInt(tdatenreihe1(px - 1) + tdatenreihe2(px - 1))
                                .DataLabel.Text = hproj.startDate.AddDays(bis - 1).ToShortDateString

                            End With

                        Next

                        .ChartType = Excel.XlChartType.xlBarStacked

                    End With


                    .HasAxis(Excel.XlAxisType.xlCategory) = True
                    .HasAxis(Excel.XlAxisType.xlValue) = True

                    With CType(.Axes(Excel.XlAxisType.xlCategory), Excel.Axis)
                        .HasTitle = False
                        .ReversePlotOrder = True
                    End With


                    With CType(.Axes(Excel.XlAxisType.xlValue), Excel.Axis)
                        .HasTitle = True
                        .MinimumScale = 0
                        .MaximumScale = CInt(maxscale / 365 * 12) + 3
                        .HasMajorGridlines = True
                        .MajorUnit = 12

                        With .AxisTitle
                            .Characters.Text = "Monate"
                            .Font.Size = awinSettings.fontsizeItems + 4
                        End With
                    End With

                    .HasLegend = False

                    .HasTitle = True
                    .ChartTitle.Text = diagramTitle
                    .ChartTitle.Font.Size = awinSettings.fontsizeTitle
                    .ChartTitle.Format.TextFrame2.TextRange.Characters(titelTeilLaengen(0) + 1, _
                                                                        titelTeilLaengen(1)).Font.Size = awinSettings.fontsizeLegend
                    .Location(Where:=XlChartLocation.xlLocationAsObject, Name:=CType(appInstance.Worksheets(arrWsNames(3)), Excel.Worksheet).Name)
                End With

                ' jetzt kommt die Korrektur der Größe; herausfinden, wieviel Raum die Axis Beschriftung einnimmt ... 
                With CType(.ChartObjects(anzDiagrams + 1), Excel.ChartObject)

                    .Chart.ChartTitle.Format.TextFrame2.TextRange.Characters(titelTeilLaengen(0) + 1, _
                                                                   titelTeilLaengen(1)).Font.Size = awinSettings.fontsizeLegend
                    .Top = top
                    .Height = (anzPhasen - 1) * 20 + 110

                    Dim axCleft As Double, axCwidth As Double
                    If CBool(.Chart.HasAxis(Excel.XlAxisType.xlCategory)) Then
                        With CType(.Chart.Axes(Excel.XlAxisType.xlCategory), Excel.Axis)
                            axCleft = .Left
                            axCwidth = .Width
                        End With

                        If left - axCwidth < 1 Then
                            .Left = 1
                            .Width = width + left + 9
                        Else
                            .Left = left - axCwidth
                            .Width = width + axCwidth + 9
                        End If

                    Else
                        .Left = left
                        .Width = width
                    End If

                    .Name = kennung


                End With

                repObj = CType(.ChartObjects(anzDiagrams + 1), Excel.ChartObject)

            End If


        End With


        'Call awinScrollintoView()
        appInstance.EnableEvents = formerEE
        'appInstance.ScreenUpdating = formerSU


    End Sub


    '
    ' Prozedur zeigt die Phasen Verläufe eines Projektes an
    ' 
    '
    ''' <summary>
    ''' aktualisiert das PhasenChart (createPhasesBalken) in der Zeit-Maschine
    ''' </summary>
    ''' <param name="hproj"></param>
    ''' <param name="chtobj"></param>
    ''' <remarks></remarks>
    Public Sub updatePhasesBalken(ByVal hproj As clsProjekt, ByRef chtobj As Excel.ChartObject, _
                                  ByVal auswahl As Integer, ByVal changeScale As Boolean)
        Dim diagramTitle As String

        Dim anzPhasen As Integer

        Dim plenInDays As Integer
        Dim i As Integer
        Dim Xdatenreihe() As String
        Dim valueColor() As Object
        Dim tdatenreihe1() As Double, mdatenreihe() As Double, tdatenreihe2() As Double, tdatenreihe3() As Double
        Dim titelTeile(1) As String
        Dim titelTeilLaengen(1) As Integer
        Dim tmpcollection As New Collection
        Dim kennung As String = " "
        Dim maxscale As Double




        Dim pname As String = hproj.name
        tmpcollection.Add(hproj.getShapeText & "#" & auswahl.ToString)
        kennung = calcChartKennung("pr", PTprdk.Phasen, tmpcollection)



        Dim formerEE As Boolean = appInstance.EnableEvents
        'Dim formerSU As Boolean = appInstance.ScreenUpdating
        appInstance.EnableEvents = False
        'appInstance.ScreenUpdating = False



        titelTeile(0) = hproj.getShapeText & " ,  " & hproj.startDate.ToShortDateString & _
                                      " - " & hproj.startDate.AddDays(hproj.dauerInDays - 1).ToShortDateString & vbLf
        titelTeile(1) = " (" & hproj.timeStamp.ToString & ") "


        titelTeilLaengen(0) = titelTeile(0).Length
        titelTeilLaengen(1) = titelTeile(1).Length

        diagramTitle = titelTeile(0) & titelTeile(1)


        '
        ' hole die Projektdauer
        '
        With hproj
            plenInDays = .dauerInDays
        End With

        '
        ' hole die Anzahl Phasen, die in diesem Projekt vorkommen
        '
        anzPhasen = hproj.Liste.Count

        ' sonst gibt es gleich 
        If anzPhasen < 1 Then
            ReDim Xdatenreihe(0)
            ReDim tdatenreihe1(0)
            ReDim mdatenreihe(0)
            ReDim tdatenreihe2(0)
            ReDim tdatenreihe3(0)
            ReDim valueColor(0)
        Else
            ReDim Xdatenreihe(anzPhasen - 1)
            ReDim tdatenreihe1(anzPhasen - 1)
            ReDim mdatenreihe(anzPhasen - 1)
            ReDim tdatenreihe2(anzPhasen - 1)
            ReDim tdatenreihe3(anzPhasen - 1)
            ReDim valueColor(anzPhasen - 1)

        End If





        'ReDim hsum(anzPhasen - 1)


        For i = 1 To anzPhasen

            With hproj.Liste.Item(i - 1)
                Xdatenreihe(i - 1) = .name
                tdatenreihe1(i - 1) = .startOffsetinDays
                tdatenreihe2(i - 1) = .dauerInDays
                tdatenreihe3(i - 1) = 0
                Try
                    valueColor(i - 1) = .Farbe
                Catch ex As Exception
                    ' dann ist es die Farbe des Projektes
                    valueColor(i - 1) = hproj.farbe
                End Try

            End With
        Next i

        'Dim dlfontsize As Double

        With chtobj.Chart


            '' remove extra series
            'Do Until .SeriesCollection.Count = 0
            '    .SeriesCollection(1).Delete()
            'Loop

            'Aufbau der Series 

            'With .SeriesCollection.NewSeries
            With .SeriesCollection(1)

                For i = 0 To anzPhasen - 1
                    mdatenreihe(i) = tdatenreihe1(i) / 365 * 12
                Next
                .Name = "null1"
                .Interior.ColorIndex = -4142
                .Values = mdatenreihe
                .XValues = Xdatenreihe
                'ur: 22.07.2014: die DataLabel bleiben wie im Chart bereits definiert.
                '.HasDataLabels = False



                For px = 1 To anzPhasen

                    With CType(.Points(px), Excel.Point)

                        If tdatenreihe1(px - 1) < 90 Then
                            .HasDataLabel = False
                        Else

                            .HasDataLabel = True
                            .DataLabel.Text = hproj.startDate.AddDays(tdatenreihe1(px - 1)).ToShortDateString

                            '.DataLabel.Format.TextFrame2.TextRange.Characters(, ).Font.Size = dlfontsize
                            If mdatenreihe(px - 1) < 5 Then

                                Try
                                    .DataLabel.Position = Excel.XlDataLabelPosition.xlLabelPositionOutsideEnd
                                    '.DataLabel.Font.Size = awinSettings.fontsizeItems
                                Catch ex As Exception
                                    .DataLabel.Position = Excel.XlDataLabelPosition.xlLabelPositionInsideEnd
                                End Try

                            Else
                                .DataLabel.Position = Excel.XlDataLabelPosition.xlLabelPositionInsideEnd
                            End If

                        End If

                    End With

                Next

                .ChartType = Excel.XlChartType.xlBarStacked
            End With


            'With .SeriesCollection.NewSeries
            With .SeriesCollection(2)

                For i = 0 To anzPhasen - 1
                    mdatenreihe(i) = tdatenreihe2(i) / 365 * 12

                    If maxscale < (tdatenreihe1(i) + tdatenreihe2(i)) / 365 * 12 Then
                        maxscale = (tdatenreihe1(i) + tdatenreihe2(i)) / 365 * 12
                    End If


                Next
                .Name = "Phasen Zeitraum"
                .Values = mdatenreihe
                .XValues = Xdatenreihe

                .HasDataLabels = True

                '.DataLabels.Font.Size = awinSettings.fontsizeItems
                .DataLabels.Position = Excel.XlDataLabelPosition.xlLabelPositionCenter

                For i = 1 To anzPhasen
                    With CType(.Points(i), Excel.Point)
                        .Interior.Color = valueColor(i - 1)

                        If mdatenreihe(i - 1) <= 3 Then
                            .DataLabel.Text = tdatenreihe2(i - 1).ToString
                        Else
                            .DataLabel.Text = tdatenreihe2(i - 1).ToString & " Tage"
                        End If
                    End With
                Next


                .ChartType = Excel.XlChartType.xlBarStacked
            End With


            'With .SeriesCollection.NewSeries
            With .SeriesCollection(3)

                .Name = "null2"
                .Interior.colorindex = -4142

                .Values = tdatenreihe3
                .XValues = Xdatenreihe

                .HasDataLabels = True

                '.DataLabels.Font.Size = awinSettings.fontsizeItems + 2
                .DataLabels.Position = Excel.XlDataLabelPosition.xlLabelPositionInsideBase


                Dim bis As Integer
                For px = 1 To anzPhasen

                    With CType(.Points(px), Excel.Point)

                        bis = CInt(tdatenreihe1(px - 1) + tdatenreihe2(px - 1))
                        .DataLabel.Text = hproj.startDate.AddDays(bis - 1).ToShortDateString

                    End With

                Next

                .ChartType = Excel.XlChartType.xlBarStacked

            End With


            If CBool(.HasAxis(Excel.XlAxisType.xlValue)) Then

                With CType(.Axes(Excel.XlAxisType.xlValue), Excel.Axis)
                    ' das ist dann relevant, wenn ein anderes Projekt selektiert wird, das über die aktuelle Skalierung 
                    ' hinausgehende Werte hat 

                    If changeScale Then
                        .MinimumScale = 0
                        If Not (.MaximumScaleIsAuto) Then

                            If maxscale > Math.Round(.MaximumScale + 5) Then
                                .MaximumScale = Math.Round(maxscale + 6)
                            End If
                            .MaximumScaleIsAuto = True
                        End If
                    End If

                    If mdatenreihe.Max > .MaximumScale - 3 Then
                        .MaximumScale = mdatenreihe.Max + 3
                    End If

                End With

            End If


            If .HasTitle Then
                .ChartTitle.Text = diagramTitle
                '.ChartTitle.Font.Size = awinSettings.fontsizeTitle
                .ChartTitle.Format.TextFrame2.TextRange.Characters(titelTeilLaengen(0) + 1, _
                       titelTeilLaengen(1)).Font.Size = awinSettings.fontsizeLegend
            End If

        End With

        chtobj.Name = kennung

        appInstance.EnableEvents = formerEE



    End Sub

    
    ''' <summary>
    ''' aktualisiert das Portfolio Einzelprojekt Chart 
    '''  
    ''' </summary>
    ''' <param name="hproj">das genau eine Projekt, das angezeigt werden soll</param>
    ''' <param name="chtobj">das Chart, das aktualisiert werden soll </param>
    ''' <param name="auswahl">gibt an, ob Projektfarbe oder AmpelFarbe angezeigt werden soll</param>
    ''' <remarks></remarks>
    Sub updateProjectPfDiagram(ByVal hproj As clsProjekt, ByRef chtobj As Excel.ChartObject, ByVal auswahl As Integer)

        Dim i As Integer
        Dim pName As String
        Dim anzBubbles As Integer
        Dim riskValues() As Double, bubbleValues() As Double, tempArray() As Double
        Dim xAchsenValues() As Double
        Dim nameValues() As String
        Dim colorValues() As Object
        Dim positionValues() As String
        Dim diagramTitle As String
        Dim showLabels As Boolean
        Dim showNegativeValues As Boolean = False
        Dim projektListe As New Collection
        Dim charttype As Integer
        Dim tmpstr(5) As String
        Dim isSingleProject As Boolean = False
        Dim tmpcollection As New Collection
        Dim kennung As String = " "
        Dim bubbleColor As Integer = 0
        Dim titelTeile(1) As String
        Dim titelTeilLaengen(1) As Integer


        isSingleProject = True
        projektListe.Add(hproj.name)

        tmpstr = chtobj.Name.Trim.Split(New Char() {CChar("#")}, 4)
        charttype = CInt(tmpstr(1))

        pName = hproj.name
        tmpcollection.Add(pName & "#" & auswahl.ToString)
        kennung = calcChartKennung("pr", charttype, tmpcollection)

        'foundDiagramm = DiagramList.getDiagramm(chtobj.Name)
        ' event. für eine Erweiterung benötigt


        ' es handelt sich garantiert nur um ein Projekt  
        Try
            ReDim riskValues(0)
            ReDim xAchsenValues(0)
            ReDim bubbleValues(0)
            ReDim nameValues(0)
            ReDim colorValues(0)
            ReDim PfChartBubbleNames(0)
            ReDim positionValues(0)
        Catch ex As Exception
            Throw New ArgumentException("Fehler in UpdatePortfolioDiagramm " & ex.Message)
        End Try


        ' neuer Typ: 8.3.14 Abhängigkeiten
        Dim activeDepIndex As Integer           ' Kennzahl: wieviel Projekte sind abhängig, wie stark strahlt das Projekt 
        Dim passiveDepIndex As Integer          ' Kennzahl: von wievielen Projekten abhängig
        Dim activeNumber As Integer             ' Kennzahl: auf wieviele Projekte strahlt es aus ?
        Dim passiveNumber As Integer            ' Kennzahl: von wievielen Projekten abhängig 

        anzBubbles = 0



        '' Änderung 6.6 : wird aktuell noch nicht unterstützt 

        'If charttype = PTpfdk.Dependencies Then
        '    Dim deleteList As New Collection
        '    For i = 1 To projektListe.Count
        '        pName = projektListe.Item(i)
        '        Try
        '            activeNumber = allDependencies.activeNumber(pName, PTdpndncyType.inhalt)
        '            passiveNumber = allDependencies.passiveNumber(pName, PTdpndncyType.inhalt)
        '            If activeNumber = 0 And passiveNumber = 0 Then
        '                deleteList.Add(pName)
        '            End If
        '        Catch ex As Exception

        '        End Try
        '    Next

        '    ' jetzt müssen die Projekte rausgenommen werden, die keine Abhängigkeiten haben 
        '    For i = 1 To deleteList.Count
        '        pName = deleteList.Item(i)
        '        Try
        '            projektListe.Remove(pName)
        '        Catch ex As Exception

        '        End Try
        '    Next
        'End If






        For i = 1 To projektListe.Count

            pName = CStr(projektListe.Item(i))

            Try

                With hproj

                    ' neuer Typ: 8.3.14 Abhängigkeiten
                    If charttype = PTpfdk.Dependencies Then
                        ' wird um eins erhöht , damit es nicht auf der Nullinie liegt 
                        activeDepIndex = allDependencies.activeIndex(pName, PTdpndncyType.inhalt) + 1
                        activeNumber = allDependencies.activeNumber(pName, PTdpndncyType.inhalt)
                        ' wird um eins erhöht , damit es nicht auf der Nullinie liegt 
                        passiveDepIndex = allDependencies.passiveIndex(pName, PTdpndncyType.inhalt) + 1
                        passiveNumber = allDependencies.passiveNumber(pName, PTdpndncyType.inhalt)
                        riskValues(anzBubbles) = activeDepIndex
                    Else
                        riskValues(anzBubbles) = .Risiko
                    End If

                    If bubbleColor = PTpfdk.ProjektFarbe Then

                        ' Projekttyp wird farblich gekennzeichent
                        colorValues(anzBubbles) = .farbe

                    Else ' bubbleColor ist AmpelFarbe

                        ' ProjektStatus wird farblich gekennzeichnet
                        Select Case .ampelStatus
                            Case 0
                                '"Ampel nicht bewertet"
                                colorValues(anzBubbles) = awinSettings.AmpelNichtBewertet
                            Case 1
                                '"Ampel Grün"
                                colorValues(anzBubbles) = awinSettings.AmpelGruen
                            Case 2
                                '"Ampel Gelb"
                                colorValues(anzBubbles) = awinSettings.AmpelGelb
                            Case 3
                                '"Ampel Rot"
                                colorValues(anzBubbles) = awinSettings.AmpelRot
                        End Select
                    End If

                    Select Case charttype
                        Case PTprdk.StrategieRisiko
                            'Strategie
                            xAchsenValues(anzBubbles) = .StrategicFit
                            bubbleValues(anzBubbles) = .ProjectMarge
                            nameValues(anzBubbles) = .name
                            PfChartBubbleNames(anzBubbles) = Format(bubbleValues(anzBubbles), "##0.#%")


                        Case PTprdk.FitRisikoVol

                            xAchsenValues(anzBubbles) = .StrategicFit
                            bubbleValues(anzBubbles) = .volume
                            nameValues(anzBubbles) = .name
                            PfChartBubbleNames(anzBubbles) = hproj.name & _
                                    " (" & Format(bubbleValues(anzBubbles) / 1000, "##0.#") & " T)"


                        Case PTprdk.ZeitRisiko

                            xAchsenValues(anzBubbles) = .dauerInDays / 365 * 12                    'Zeit
                            bubbleValues(anzBubbles) = System.Math.Round(.volume / 10000) * 10
                            'tmpstr = .name.Split(New Char() {" "}, 10)                             'Zeit/Risiko
                            'nameValues(anzBubbles) = tmpstr(0) & " (" & Format(bubbleValues(anzBubbles), "##0.#") & " T)" 
                            nameValues(anzBubbles) = .name & " (" & Format(bubbleValues(anzBubbles), "##0.#") & " T)"
                            PfChartBubbleNames(anzBubbles) = .name & _
                                    " (" & Format(bubbleValues(anzBubbles), "##0.#") & " T)"

                        Case PTprdk.ComplexRisiko

                            xAchsenValues(anzBubbles) = .complexity                                'Complex
                            bubbleValues(anzBubbles) = .volume                                     'Bubblegröße gemäß Volumen
                            nameValues(anzBubbles) = .name
                            PfChartBubbleNames(anzBubbles) = hproj.name & _
                             " (" & Format(bubbleValues(anzBubbles) / 1000, "##0.#") & " T)"


                        Case PTprdk.Dependencies
                            ' neuer Typ: 8.3.14 Abhängigkeiten

                            xAchsenValues(anzBubbles) = passiveDepIndex                            'Abhängigkeiten
                            bubbleValues(anzBubbles) = .StrategicFit
                            nameValues(anzBubbles) = .name

                            PfChartBubbleNames(anzBubbles) = .name & _
                                " (" & passiveNumber.ToString & ", " & activeNumber.ToString & ")"

                    End Select
                End With
                anzBubbles = anzBubbles + 1
            Catch ex As Exception

            End Try
        Next

        Select Case charttype
            Case PTprdk.StrategieRisiko

                titelTeile(0) = summentitel2 & " " & hproj.name & vbLf
                titelTeile(1) = " (" & hproj.timeStamp.ToString & ") "

            Case PTpfdk.FitRisikoVol

                titelTeile(0) = portfolioDiagrammtitel(PTprdk.FitRisikoVol) & " " & hproj.name & vbLf
                titelTeile(1) = " (" & hproj.timeStamp.ToString & ") "

            Case PTpfdk.ZeitRisiko

                titelTeile(0) = portfolioDiagrammtitel(PTprdk.ZeitRisiko) & " " & hproj.name & vbLf
                titelTeile(1) = " (" & hproj.timeStamp.ToString & ") "

            Case PTpfdk.ComplexRisiko

                titelTeile(0) = portfolioDiagrammtitel(PTprdk.ComplexRisiko) & " " & hproj.name & vbLf
                titelTeile(1) = " (" & hproj.timeStamp.ToString & ") "

            Case PTpfdk.Dependencies
                ' neuer Typ: 8.3.14 Abhängigkeiten

                titelTeile(0) = portfolioDiagrammtitel(PTprdk.Dependencies) & " " & hproj.name & vbLf
                titelTeile(1) = " (" & hproj.timeStamp.ToString & ") "

            Case Else
                diagramTitle = "Chart-Typ existiert nicht"
        End Select

        titelTeilLaengen(0) = titelTeile(0).Length
        titelTeilLaengen(1) = titelTeile(1).Length

        diagramTitle = titelTeile(0) & titelTeile(1)


        ' bestimmen der besten Position für die Werte ...
        Dim labelPosition(4) As String
        labelPosition(0) = "oben"
        labelPosition(1) = "rechts"
        labelPosition(2) = "unten"
        labelPosition(3) = "links"
        labelPosition(4) = "mittig"

        For i = 0 To anzBubbles - 1

            positionValues(i) = pfchartIstFrei(i, xAchsenValues, riskValues)

        Next


        ReDim tempArray(anzBubbles - 1)

        Dim formerEE As Boolean = appInstance.EnableEvents
        appInstance.EnableEvents = False

        ' nur dann neue Series-Collection aufbauen, wenn auch tatsächlich was in der Projektliste ist ..

        If projektListe.Count > 0 Then

            With CType(chtobj.Chart, Excel.Chart)

                showLabels = True

                ' Einstellungen der vorhandenen SeriesCollection merken

                Dim pts As Excel.Points = CType(.SeriesCollection(1).Points, Excel.Points)
                Dim dlFontSize As Double
                Dim dlFontBackground As Double
                Dim dlFontBold As Boolean
                Dim dlFontColorIndex As Integer
                Dim dlFontColor As Integer
                Dim dlFontFontStyle As String = ""
                Dim dlFontItalic As Boolean
                Dim dlFontStrikethrough As Boolean
                Dim dlFontSuperscript As Boolean
                Dim dlFontSubscript As Boolean
                Dim dlFontUnderline As Double

                For i = 1 To pts.Count

                    With CType(.SeriesCollection(1).Points(i), Excel.Point)

                        Try
                            If .HasDataLabel = True Then

                                With .DataLabel
                                    dlFontSize = CDbl(.Font.Size)
                                    dlFontBackground = CDbl(.Font.Background)
                                    dlFontBold = CBool(.Font.Bold)
                                    dlFontColorIndex = CInt(.Font.ColorIndex)
                                    dlFontFontStyle = CStr(.Font.FontStyle)
                                    dlFontItalic = CBool(.Font.Italic)
                                    dlFontStrikethrough = CBool(.Font.Strikethrough)
                                    dlFontSubscript = CBool(.Font.Subscript)
                                    dlFontSuperscript = CBool(.Font.Superscript)
                                    'ur 21.07.2015 dlFontUnderline = CDbl(.Font.Underline)

                                    dlFontSize = CDbl(.Font.Size)
                                    If .Font.Color <> awinSettings.AmpelRot Then
                                        dlFontColor = CInt(.Font.Color)
                                    End If

                                End With
                            End If

                        Catch ex As Exception

                        End Try
                    End With

                Next i



                ' remove old series
                Do Until .SeriesCollection.Count = 0
                    .SeriesCollection(1).Delete()
                Loop



                .SeriesCollection.NewSeries()
                .SeriesCollection(1).name = diagramTitle
                .SeriesCollection(1).ChartType = Excel.XlChartType.xlBubble3DEffect

                For i = 1 To anzBubbles
                    tempArray(i - 1) = xAchsenValues(i - 1)
                Next i
                .SeriesCollection(1).XValues = tempArray ' strategic

                For i = 1 To anzBubbles
                    tempArray(i - 1) = riskValues(i - 1)
                Next i
                .SeriesCollection(1).Values = tempArray

                For i = 1 To anzBubbles
                    If bubbleValues(i - 1) < 0.01 And bubbleValues(i - 1) > -0.01 Then
                        tempArray(i - 1) = 0.01
                    ElseIf bubbleValues(i - 1) < 0 Then
                        ' negative Werte werden Positiv dargestellt mit roten Beschriftung siehe unten
                        tempArray(i - 1) = System.Math.Abs(bubbleValues(i - 1))
                    Else
                        tempArray(i - 1) = bubbleValues(i - 1)
                    End If
                Next i


                .SeriesCollection(1).BubbleSizes = tempArray

                Dim series1 As Excel.Series = _
                        CType(.SeriesCollection(1),  _
                                Excel.Series)
                Dim point1 As Excel.Point = _
                            CType(series1.Points(1), Excel.Point)

                'Dim testName As String

                Dim bubblePoint As Excel.Point
                For i = 1 To anzBubbles

                    bubblePoint = CType(.SeriesCollection(1).Points(i), Excel.Point)

                    With CType(.SeriesCollection(1).Points(i), Excel.Point)

                        If showLabels Then
                            Try
                                .HasDataLabel = True

                                With .DataLabel
                                    .Text = PfChartBubbleNames(i - 1)
                                    .Font.Size = dlFontSize
                                    .Font.Background = dlFontBackground
                                    .Font.Bold = dlFontBold
                                    .Font.Color = dlFontColor
                                    .Font.ColorIndex = dlFontColorIndex
                                    .Font.FontStyle = dlFontFontStyle
                                    .Font.Italic = dlFontItalic
                                    .Font.Strikethrough = dlFontStrikethrough
                                    .Font.Subscript = dlFontSubscript
                                    .Font.Superscript = dlFontSuperscript
                                    'ur: 21.07.2015: .Font.Underline = dlFontUnderline

                                    'ur: 17.7.2014: fontsize kommt vom existierenden Chart
                                    '.Font.Size = awinSettings.CPfontsizeItems

                                    ' bei negativen Werten erfolgt die Beschriftung in roter Farbe  ..

                                    If bubbleValues(i - 1) < 0 Then
                                        .Font.Color = awinSettings.AmpelRot

                                    End If
                                    Select Case positionValues(i - 1)
                                        Case labelPosition(0)
                                            .Position = Excel.XlDataLabelPosition.xlLabelPositionAbove
                                        Case labelPosition(1)
                                            .Position = Excel.XlDataLabelPosition.xlLabelPositionRight
                                        Case labelPosition(2)
                                            .Position = Excel.XlDataLabelPosition.xlLabelPositionBelow
                                        Case labelPosition(3)
                                            .Position = Excel.XlDataLabelPosition.xlLabelPositionLeft
                                        Case Else
                                            .Position = Excel.XlDataLabelPosition.xlLabelPositionCenter
                                    End Select
                                End With
                            Catch ex As Exception

                            End Try
                        Else
                            .HasDataLabel = False
                        End If

                        .Interior.Color = colorValues(i - 1)

                        ' bei negativen Werten erfolgt die Beschriftung in roter Farbe  ..
                        If bubbleValues(i - 1) < 0 Then
                            .DataLabel.Font.Color = awinSettings.AmpelRot
                        End If
                    End With
                Next i



                '.ChartGroups(1).BubbleScale = sollte in Abhängigkeit der width gemacht werden 
                With .ChartGroups(1)

                    .BubbleScale = 20
                    .SizeRepresents = Microsoft.Office.Interop.Excel.XlSizeRepresents.xlSizeIsArea

                    If showNegativeValues Then
                        .shownegativeBubbles = True
                    Else
                        .shownegativeBubbles = False
                    End If
                End With


                If .HasTitle Then
                    .ChartTitle.Text = diagramTitle
                    ' ur: 21.07.2014 für Chart-Cockpit auskommentiert
                    '.ChartTitle.Font.Size = awinSettings.fontsizeTitle
                    '.ChartTitle.Format.TextFrame2.TextRange.Characters(titelTeilLaengen(0) + 1, _
                    '   titelTeilLaengen(1)).Font.Size = awinSettings.fontsizeLegend
                End If


            End With

        End If

        chtobj.Name = kennung
        appInstance.EnableEvents = formerEE




    End Sub


    

    ''' <summary>
    ''' zeigt Soll-/Ist zu Personalkosten, Sonstige Kosten , Gesamtkosten an 
    ''' das wird gesteuert über auswahl ; damit können auch beliebige andere angezeigt werden 
    ''' muss dann aber noch implementiert werden 
    ''' </summary>
    ''' <param name="hproj"></param>
    ''' <param name="reportObj"></param>
    ''' <param name="heute"></param>
    ''' <param name="auswahl"></param>
    ''' <param name="top"></param>
    ''' <param name="left"></param>
    ''' <param name="height"></param>
    ''' <param name="width"></param>
    ''' <remarks></remarks>
    Sub createSollIstOfProject(ByRef hproj As clsProjekt, ByRef reportObj As Excel.ChartObject, ByVal heute As Date, ByVal auswahl As Integer, ByVal qualifier As String, ByVal vglBaseline As Boolean, _
                                   ByVal top As Double, ByVal left As Double, ByVal height As Double, ByVal width As Double)
        Dim chtobj As Excel.ChartObject
        Dim anzDiagrams As Integer
        Dim i As Integer, ix As Integer = 0
        Dim found As Boolean
        Dim abbruch As Boolean = False
        Dim pname As String = hproj.name
        Dim fullname As String = hproj.getShapeText
        Dim kennung As String = " "
        Dim diagramTitle As String = " "
        Dim zE As String = "(" & awinSettings.kapaEinheit & ")"
        Dim titelTeile(2) As String
        Dim titelTeilLaengen(2) As Integer
        Dim kontrollWert As Double
        Dim vgl As Date

        Dim isMinMax As Boolean = False

        Dim beauftragung As clsProjekt
        Dim lastPlan As clsProjekt
        Dim anzSnapshots As Integer = projekthistorie.Count


        Dim formerEE As Boolean = appInstance.EnableEvents
        appInstance.EnableEvents = False


        ' Änderung 18.6 : Unterscheidung zwischen Soll-/Ist Vergleichen und Min/Max Vergleichen 

        If hproj.Status <> ProjektStatus(0) Then
            ' Soll-Ist Vergleich
            isMinMax = False

            Try
                beauftragung = projekthistorie.beauftragung
            Catch ex As Exception
                Throw New ArgumentException("es gibt keine Beauftragung")
            End Try


            ' finde in der Projekt-Historie das Projekt, das direkt vor hproj gespeichert wurde
            ' 
            vgl = hproj.timeStamp.AddMinutes(-1)
            lastPlan = projekthistorie.ElementAtorBefore(vgl)

        Else
            ' Min-Max Vergleich 
            isMinMax = True
            Dim minIndex As Integer = 0
            Dim maxIndex As Integer = 0
            Dim minValue As Double, maxValue As Double
            Dim tmpValue As Double
            Select Case auswahl
                Case 1
                    ' Personalkosten
                    With projekthistorie.ElementAt(0)
                        minValue = .getAllPersonalKosten.Sum
                        maxValue = .getAllPersonalKosten.Sum
                    End With

                    For s = 1 To anzSnapshots - 1
                        With projekthistorie.ElementAt(s)
                            tmpValue = .getAllPersonalKosten.Sum
                            If tmpValue < minValue Then
                                minIndex = s
                                minValue = tmpValue
                            End If
                            If tmpValue > maxValue Then
                                maxIndex = s
                                maxValue = tmpValue
                            End If
                        End With
                    Next


                Case 2
                    ' Sonstige Kosten
                    With projekthistorie.ElementAt(0)
                        minValue = .getGesamtAndereKosten.Sum
                        maxValue = .getGesamtAndereKosten.Sum
                    End With

                    For s = 1 To anzSnapshots - 1
                        With projekthistorie.ElementAt(s)
                            tmpValue = .getGesamtAndereKosten.Sum
                            If tmpValue < minValue Then
                                minIndex = s
                                minValue = tmpValue
                            End If
                            If tmpValue > maxValue Then
                                maxIndex = s
                                maxValue = tmpValue
                            End If
                        End With
                    Next
                Case 3
                    ' Gesamtkosten
                    With projekthistorie.ElementAt(0)
                        minValue = .getGesamtKostenBedarf.Sum
                        maxValue = .getGesamtKostenBedarf.Sum
                    End With

                    For s = 1 To anzSnapshots - 1
                        With projekthistorie.ElementAt(s)
                            tmpValue = .getGesamtKostenBedarf.Sum
                            If tmpValue < minValue Then
                                minIndex = s
                                minValue = tmpValue
                            End If
                            If tmpValue > maxValue Then
                                maxIndex = s
                                maxValue = tmpValue
                            End If
                        End With
                    Next
                Case 4
                    ' Rollen mit Qualifier
                    With projekthistorie.ElementAt(0)
                        minValue = .getPersonalKosten(qualifier).Sum
                        maxValue = .getPersonalKosten(qualifier).Sum
                    End With

                    For s = 1 To anzSnapshots - 1
                        With projekthistorie.ElementAt(s)
                            tmpValue = .getPersonalKosten(qualifier).Sum
                            If tmpValue < minValue Then
                                minIndex = s
                                minValue = tmpValue
                            End If
                            If tmpValue > maxValue Then
                                maxIndex = s
                                maxValue = tmpValue
                            End If
                        End With
                    Next


                Case 5
                    ' Kostenart mit Qualifier
                    With projekthistorie.ElementAt(0)
                        minValue = .getKostenBedarf(qualifier).Sum
                        maxValue = .getKostenBedarf(qualifier).Sum
                    End With

                    For s = 1 To anzSnapshots - 1
                        With projekthistorie.ElementAt(s)
                            tmpValue = .getKostenBedarf(qualifier).Sum
                            If tmpValue < minValue Then
                                minIndex = s
                                minValue = tmpValue
                            End If
                            If tmpValue > maxValue Then
                                maxIndex = s
                                maxValue = tmpValue
                            End If
                        End With
                    Next

                Case Else
                    ' Gesamtkosten
                    With projekthistorie.ElementAt(0)
                        minValue = .getGesamtKostenBedarf.Sum
                        maxValue = .getGesamtKostenBedarf.Sum
                    End With

                    For s = 1 To anzSnapshots - 1
                        With projekthistorie.ElementAt(s)
                            tmpValue = .getGesamtKostenBedarf.Sum
                            If tmpValue < minValue Then
                                minIndex = s
                                minValue = tmpValue
                            End If
                            If tmpValue > maxValue Then
                                maxIndex = s
                                maxValue = tmpValue
                            End If
                        End With
                    Next
            End Select

            Try
                beauftragung = projekthistorie.ElementAt(minIndex)
                lastPlan = projekthistorie.ElementAt(maxIndex)
            Catch ex As Exception
                Throw New ArgumentException("Fehler in Min-/Max Bestimmung " & ex.Message)
            End Try

        End If



        ' Ende Ergänzung 18.6 Min/Max 





        Dim minColumn As Integer, maxColumn As Integer, heuteColumn As Integer = getColumnOfDate(heute)
        Dim pastAndFuture As Boolean = False
        Dim future As Boolean = True

        Dim werteB(beauftragung.anzahlRasterElemente - 1) As Double
        Dim werteL(lastPlan.anzahlRasterElemente - 1) As Double
        Dim werteC(hproj.anzahlRasterElemente - 1) As Double

        Dim Xdatenreihe() As String
        Dim tdatenreiheB() As Double
        Dim tdatenreiheL() As Double
        Dim tdatenreiheC() As Double

        ' Bestimmen der Werte 
        Select Case auswahl
            Case 1
                ' Personalkosten
                If isMinMax Then
                    titelTeile(0) = "Min/Max Personalkosten (T€)" & vbLf
                Else
                    titelTeile(0) = "Soll/Ist Personalkosten (T€)" & vbLf
                End If

                kennung = "Soll/Ist Personalkosten"
                werteB = beauftragung.getAllPersonalKosten
                werteL = lastPlan.getAllPersonalKosten
                werteC = hproj.getAllPersonalKosten
            Case 2
                ' Sonstige Kosten
                If isMinMax Then
                    titelTeile(0) = "Min/Max Sonstige Kosten (T€)" & vbLf
                Else
                    titelTeile(0) = "Soll/Ist Sonstige Kosten (T€)" & vbLf
                End If

                kennung = "Soll/Ist Sonstige Kosten"
                werteB = beauftragung.getGesamtAndereKosten
                werteL = lastPlan.getGesamtAndereKosten
                werteC = hproj.getGesamtAndereKosten

            Case 3
                ' Gesamt Kosten
                If isMinMax Then
                    titelTeile(0) = "Min/Max Gesamtkosten (T€)" & vbLf
                Else
                    titelTeile(0) = "Soll/Ist Gesamtkosten (T€)" & vbLf
                End If

                kennung = "Soll/Ist Gesamtkosten"
                werteB = beauftragung.getGesamtKostenBedarf
                werteL = lastPlan.getGesamtKostenBedarf
                werteC = hproj.getGesamtKostenBedarf
            Case 4
                ' Rollen mit Qualifier
                If isMinMax Then
                    titelTeile(0) = "Min/Max " & qualifier & "(" & awinSettings.kapaEinheit & ")" & vbLf
                Else
                    titelTeile(0) = "Soll-/Ist " & qualifier & "(" & awinSettings.kapaEinheit & ")" & vbLf
                End If

                kennung = "Rolle " & qualifier
                Try
                    werteB = beauftragung.getRessourcenBedarf(qualifier)
                    werteL = lastPlan.getRessourcenBedarf(qualifier)
                    werteC = hproj.getRessourcenBedarf(qualifier)
                Catch ex As Exception
                    Throw New ArgumentException(ex.Message & vbLf & qualifier & " nicht gefunden")
                End Try

            Case 5
                ' Kostenart mit Qualifier
                If isMinMax Then
                    titelTeile(0) = "Min/Max " & qualifier & " (T€)" & vbLf
                Else
                    titelTeile(0) = "Soll-/Ist " & qualifier & " (T€)" & vbLf
                End If

                kennung = "Kostenart " & qualifier
                Try
                    werteB = beauftragung.getKostenBedarf(qualifier)
                    werteL = lastPlan.getKostenBedarf(qualifier)
                    werteC = hproj.getKostenBedarf(qualifier)
                Catch ex As Exception
                    Throw New ArgumentException(ex.Message & vbLf & qualifier & " nicht gefunden")
                End Try

            Case Else
                ' Gesamt Kosten
                If isMinMax Then
                    titelTeile(0) = "Min/Max Gesamtkosten (T€)" & vbLf
                Else
                    titelTeile(0) = "Soll/Ist Gesamtkosten (T€)" & vbLf
                End If

                kennung = "Soll/Ist Gesamtkosten"
                werteB = beauftragung.getGesamtKostenBedarf
                werteL = lastPlan.getGesamtKostenBedarf
                werteC = hproj.getGesamtKostenBedarf
                auswahl = 3

        End Select

        titelTeilLaengen(0) = titelTeile(0).Length
        titelTeile(1) = fullname & vbLf
        titelTeilLaengen(1) = titelTeile(1).Length
        titelTeile(2) = " (" & hproj.timeStamp.ToString & ") "
        titelTeilLaengen(2) = titelTeile(2).Length
        diagramTitle = titelTeile(0) & titelTeile(1) & titelTeile(2)

        minColumn = 10000
        If beauftragung.Start < minColumn Then
            minColumn = beauftragung.Start
        End If

        If lastPlan.Start < minColumn Then
            minColumn = lastPlan.Start
        End If

        If hproj.Start < minColumn Then
            minColumn = hproj.Start
        End If


        With hproj
            maxColumn = .Start + .anzahlRasterElemente - 1
        End With

        With beauftragung
            If maxColumn < .Start + .anzahlRasterElemente - 1 Then
                maxColumn = .Start + .anzahlRasterElemente - 1
            End If
        End With

        With lastPlan
            If maxColumn < .Start + .anzahlRasterElemente - 1 Then
                maxColumn = .Start + .anzahlRasterElemente - 1
            End If
        End With

        If heuteColumn >= minColumn + 1 And _
            heuteColumn <= maxColumn Then

            pastAndFuture = True

        End If


        ReDim Xdatenreihe(1)
        ReDim tdatenreiheB(1)
        ReDim tdatenreiheL(1)
        ReDim tdatenreiheC(1)


        If heuteColumn >= minColumn + 1 And heuteColumn <= maxColumn Then
            Xdatenreihe(0) = "Soll/Ist-Werte" & vbLf & textZeitraum(minColumn, heuteColumn - 1)
            Xdatenreihe(1) = "Prognose" & vbLf & textZeitraum(heuteColumn, maxColumn)
        ElseIf heuteColumn > maxColumn Then
            future = False
            Xdatenreihe(0) = "Soll/Ist-Werte" & vbLf & textZeitraum(minColumn, maxColumn)
            Xdatenreihe(1) = "Prognose" & vbLf & "existiert nicht"
        ElseIf heuteColumn <= minColumn Then
            future = True
            Xdatenreihe(0) = "Soll/Ist-Werte" & vbLf & "existieren nicht"
            Xdatenreihe(1) = "Prognose" & vbLf & textZeitraum(minColumn, maxColumn)
        End If


        Dim hsum As Double = 0.0
        ix = 0
        Dim endeIX As Integer
        With beauftragung

            'If werteB.Sum >= 100 Then
            '    kontrollWert = Math.Round(werteB.Sum)
            'Else
            kontrollWert = Math.Round(werteB.Sum)
            'End If

            endeIX = System.Math.Min(heuteColumn - 1, .Start + .anzahlRasterElemente - 1)
            For i = .Start To endeIX
                hsum = hsum + werteB(ix)
                ix = ix + 1
            Next
            'If hsum >= 100 Then
            '    tdatenreiheB(0) = Math.Round(hsum / 10) * 10
            'Else
            tdatenreiheB(0) = Math.Round(hsum)
            'End If

            tdatenreiheB(1) = kontrollWert - tdatenreiheB(0)

        End With

        ix = 0
        With lastPlan

            'If werteL.Sum >= 100 Then
            '    kontrollWert = Math.Round(werteL.Sum / 10) * 10
            'Else
            kontrollWert = Math.Round(werteL.Sum)
            'End If


            hsum = 0.0
            endeIX = System.Math.Min(heuteColumn - 1, .Start + .anzahlRasterElemente - 1)
            For i = .Start To endeIX
                hsum = hsum + werteL(ix)
                ix = ix + 1
            Next

            'If hsum >= 100 Then
            '    tdatenreiheL(0) = Math.Round(hsum / 10) * 10
            'Else
            tdatenreiheL(0) = Math.Round(hsum)
            'End If

            tdatenreiheL(1) = kontrollWert - tdatenreiheL(0)

        End With

        ix = 0
        With hproj

            'If werteC.Sum >= 100 Then
            '    kontrollWert = Math.Round(werteC.Sum / 10) * 10
            'Else
            kontrollWert = Math.Round(werteC.Sum)
            'End If

            hsum = 0.0
            endeIX = System.Math.Min(heuteColumn - 1, .Start + .anzahlRasterElemente - 1)
            For i = .Start To endeIX
                hsum = hsum + werteC(ix)
                ix = ix + 1
            Next

            'If hsum >= 100 Then
            '    tdatenreiheC(0) = Math.Round(hsum / 10) * 10
            'Else
            tdatenreiheC(0) = Math.Round(hsum)
            'End If

            tdatenreiheC(1) = kontrollWert - tdatenreiheC(0)

        End With



        With CType(appInstance.Worksheets(arrWsNames(3)), Excel.Worksheet)
            anzDiagrams = CType(.ChartObjects, Excel.ChartObjects).Count
            '
            ' um welches Diagramm handelt es sich ...
            '
            i = 1
            found = False
            While i <= anzDiagrams And Not found
                Dim chtTitle As String
                Try
                    chtTitle = CType(.ChartObjects(i), Excel.ChartObject).Chart.ChartTitle.Text
                Catch ex As Exception
                    chtTitle = " "
                End Try

                If chtTitle = diagramTitle Then
                    found = True

                Else
                    i = i + 1
                End If

            End While

            If found Then
                'Call MsgBox("Chart wird bereits angezeigt ...")
                reportObj = CType(.ChartObjects(i), Excel.ChartObject)
                appInstance.EnableEvents = formerEE
                'appInstance.ScreenUpdating = formerSU
                Exit Sub
            Else
                With appInstance.Charts.Add
                    ' remove extra series
                    Do Until .SeriesCollection.Count = 0
                        .SeriesCollection(1).Delete()
                    Loop

                    .HasAxis(Excel.XlAxisType.xlCategory) = True
                    .HasAxis(Excel.XlAxisType.xlValue) = True

                    'Dim ax As Excel.Axis

                    'With ax
                    '    .MajorUnit = 10000
                    'End With

                    With .Axes(Excel.XlAxisType.xlCategory)
                        .HasTitle = False
                        '.MinimumScale = 0
                        'With .AxisTitle
                        '    .Characters.text = "Monate"
                        '    .Font.Size = 8
                        'End With
                    End With

                    With .Axes(Excel.XlAxisType.xlValue)
                        .HasTitle = False
                        .HasMajorGridlines = False
                        .hasminorgridlines = False
                        '.MajorUnit = 10000
                        '.MinorUnit = 10000
                        '.MaximumScale = maxscale
                        '.MinimumScale = 0

                        'With .AxisTitle
                        '    .Characters.text = "Kosten"
                        '    .Font.Size = 8
                        'End With
                    End With

                    .HasLegend = True
                    With .Legend
                        .Position = Excel.Constants.xlTop
                        .Font.Size = awinSettings.fontsizeLegend
                    End With
                    .HasTitle = True
                    .ChartTitle.Text = diagramTitle
                    .ChartTitle.Font.Size = awinSettings.fontsizeTitle
                    .Location(Where:=XlChartLocation.xlLocationAsObject, Name:=appInstance.Worksheets(arrWsNames(3)).name)

                End With

                chtobj = CType(.ChartObjects(anzDiagrams + 1), Excel.ChartObject)
                chtobj.Name = fullname & "#" & kennung & "#" & "1"


            End If

            With chtobj.Chart

                .ChartTitle.Format.TextFrame2.TextRange.Characters(titelTeilLaengen(0) + _
                                                                   titelTeilLaengen(1) + 1, titelTeilLaengen(2)).Font.Size = awinSettings.fontsizeLegend

                'series

                If isMinMax Or vglBaseline Then

                    With .SeriesCollection.NewSeries
                        If isMinMax Then
                            .name = "Minimum (" & beauftragung.timeStamp.ToString("d") & ")"
                        Else
                            '.name = "Baseline (" & beauftragung.timeStamp.ToString("d") & ")"
                            .name = "Soll"
                        End If

                        '.name = "Baseline"
                        .Interior.color = awinSettings.SollIstFarbeB
                        .Values = tdatenreiheB
                        .XValues = Xdatenreihe
                        .ChartType = Excel.XlChartType.xlColumnClustered

                        If pastAndFuture Then
                            For i = 0 To 1
                                With .Points(i + 1)

                                    .HasDataLabel = True
                                    .DataLabel.text = Format(tdatenreiheB(i), "###,###0")
                                    .DataLabel.Font.Size = awinSettings.fontsizeItems

                                End With
                            Next
                        ElseIf future Then
                            With .Points(2)

                                .HasDataLabel = True
                                .DataLabel.text = Format(tdatenreiheB(1), "###,###0")
                                .DataLabel.Font.Size = awinSettings.fontsizeItems

                            End With
                        Else
                            With .Points(1)

                                .HasDataLabel = True
                                .DataLabel.text = Format(tdatenreiheB(0), "###,###0")
                                .DataLabel.Font.Size = awinSettings.fontsizeItems

                            End With
                        End If


                    End With

                End If


                If isMinMax Or Not vglBaseline Then
                    With .SeriesCollection.NewSeries
                        If isMinMax Then
                            .name = "Maximum (" & lastPlan.timeStamp.ToString("d") & ")"
                        Else
                            '.name = "Last (" & lastPlan.timeStamp.ToString("d") & ")"
                            .name = "Last"
                        End If

                        .Interior.color = awinSettings.SollIstFarbeL
                        .Values = tdatenreiheL
                        .XValues = Xdatenreihe
                        .ChartType = Excel.XlChartType.xlColumnClustered

                        If pastAndFuture Then
                            For i = 0 To 1
                                With .Points(i + 1)

                                    .HasDataLabel = True
                                    .DataLabel.text = Format(tdatenreiheL(i), "###,###0")
                                    .DataLabel.Font.Size = awinSettings.fontsizeItems

                                End With
                            Next
                        ElseIf future Then
                            With .Points(2)

                                .HasDataLabel = True
                                .DataLabel.text = Format(tdatenreiheL(1), "###,###0")
                                .DataLabel.Font.Size = awinSettings.fontsizeItems

                            End With
                        Else
                            With .Points(1)

                                .HasDataLabel = True
                                .DataLabel.text = Format(tdatenreiheL(0), "###,###0")
                                .DataLabel.Font.Size = awinSettings.fontsizeItems

                            End With
                        End If

                    End With

                End If


                With .SeriesCollection.NewSeries
                    '.name = "Current (" & hproj.timeStamp.ToString("d") & ")"
                    .name = "Ist"
                    '.name = "Current"
                    .Interior.color = awinSettings.SollIstFarbeC
                    .Values = tdatenreiheC
                    .XValues = Xdatenreihe
                    .ChartType = Excel.XlChartType.xlColumnClustered

                    If pastAndFuture Then
                        For i = 0 To 1
                            With .Points(i + 1)

                                .HasDataLabel = True
                                .DataLabel.text = Format(tdatenreiheC(i), "###,###0")
                                .DataLabel.Font.Size = awinSettings.fontsizeItems

                            End With
                        Next
                    ElseIf future Then
                        With .Points(2)

                            .HasDataLabel = True
                            .DataLabel.text = Format(tdatenreiheC(1), "###,###0")
                            .DataLabel.Font.Size = awinSettings.fontsizeItems

                        End With
                    Else
                        With .Points(1)

                            .HasDataLabel = True
                            .DataLabel.text = Format(tdatenreiheC(0), "###,###0")
                            .DataLabel.Font.Size = awinSettings.fontsizeItems

                        End With
                    End If

                End With
                .ChartGroups(1).Overlap = -50
                .ChartGroups(1).GapWidth = 150
            End With


        End With

        With chtobj
            .Top = top
            .Left = left
            .Height = height
            .Width = width
        End With

        appInstance.EnableEvents = formerEE
        reportObj = chtobj

    End Sub

    ''' <summary>
    ''' zeigt bei beauftragten Projekten den Soll-/Ist Vergleich an
    ''' zeigt bei nicht beauftragten Projekten den Vergleich zwischen Min/Max/Current an
    ''' Vorbedingung: es gibt eine Historie 
    ''' </summary>
    ''' <param name="hproj"></param>
    ''' <param name="reportObj"></param>
    ''' <param name="heute"></param>
    ''' <param name="auswahl"></param>
    ''' <param name="top"></param>
    ''' <param name="left"></param>
    ''' <param name="height"></param>
    ''' <param name="width"></param>
    ''' <remarks></remarks>
    Sub createSollIstCurveOfProject(ByRef hproj As clsProjekt, ByRef reportObj As Excel.ChartObject, ByVal heute As Date, ByVal auswahl As Integer, ByVal qualifier As String, ByVal vglBaseline As Boolean, _
                                           ByVal top As Double, ByVal left As Double, ByVal height As Double, ByVal width As Double)
        Dim chtobj As Excel.ChartObject
        Dim anzDiagrams As Integer
        Dim i As Integer, ix As Integer = 0
        Dim found As Boolean
        Dim abbruch As Boolean = False
        Dim pname As String = hproj.name
        Dim fullname As String = hproj.getShapeText
        Dim kennung As String = " "
        Dim diagramTitle As String = " "
        Dim zE As String = "(" & awinSettings.kapaEinheit & ")"
        Dim titelTeile(2) As String
        Dim titelTeilLaengen(2) As Integer
        Dim sumB As Double, sumL As Double, sumC As Double
        Dim isMinMax As Boolean = False

        Dim beauftragung As clsProjekt
        Dim lastPlan As clsProjekt
        Dim anzSnapshots As Integer = projekthistorie.Count

        If hproj.Status <> ProjektStatus(0) Then
            ' Soll-Ist Vergleich
            isMinMax = False

            Try
                beauftragung = projekthistorie.beauftragung
            Catch ex As Exception
                Throw New ArgumentException("es gibt keine Beauftragung")
            End Try

            abbruch = False
            Dim index As Integer = 0



            ' finde in der Projekt-Historie das Projekt, das direkt vor hproj gespeichert wurde

            Dim vgl As Date = hproj.timeStamp.AddMinutes(-1)

            lastPlan = projekthistorie.ElementAtorBefore(vgl)

            If IsNothing(lastPlan) Then
                Throw New ArgumentException("es gibt keinen Stand vorher")
            End If


        Else
            ' Min-Max Vergleich 
            isMinMax = True
            Dim minIndex As Integer = 0
            Dim maxIndex As Integer = 0
            Dim minValue As Double, maxValue As Double
            Dim tmpValue As Double
            Select Case auswahl
                Case 1
                    ' Personalkosten
                    With projekthistorie.ElementAt(0)
                        minValue = .getAllPersonalKosten.Sum
                        maxValue = .getAllPersonalKosten.Sum
                    End With

                    For s = 1 To anzSnapshots - 1
                        With projekthistorie.ElementAt(s)
                            tmpValue = .getAllPersonalKosten.Sum
                            If tmpValue < minValue Then
                                minIndex = s
                                minValue = tmpValue
                            End If
                            If tmpValue > maxValue Then
                                maxIndex = s
                                maxValue = tmpValue
                            End If
                        End With
                    Next


                Case 2
                    ' Sonstige Kosten
                    With projekthistorie.ElementAt(0)
                        minValue = .getGesamtAndereKosten.Sum
                        maxValue = .getGesamtAndereKosten.Sum
                    End With

                    For s = 1 To anzSnapshots - 1
                        With projekthistorie.ElementAt(s)
                            tmpValue = .getGesamtAndereKosten.Sum
                            If tmpValue < minValue Then
                                minIndex = s
                                minValue = tmpValue
                            End If
                            If tmpValue > maxValue Then
                                maxIndex = s
                                maxValue = tmpValue
                            End If
                        End With
                    Next
                Case 3
                    ' Gesamtkosten
                    With projekthistorie.ElementAt(0)
                        minValue = .getGesamtKostenBedarf.Sum
                        maxValue = .getGesamtKostenBedarf.Sum
                    End With

                    For s = 1 To anzSnapshots - 1
                        With projekthistorie.ElementAt(s)
                            tmpValue = .getGesamtKostenBedarf.Sum
                            If tmpValue < minValue Then
                                minIndex = s
                                minValue = tmpValue
                            End If
                            If tmpValue > maxValue Then
                                maxIndex = s
                                maxValue = tmpValue
                            End If
                        End With
                    Next
                Case 4
                    ' Rollen mit Qualifier
                    With projekthistorie.ElementAt(0)
                        minValue = .getPersonalKosten(qualifier).Sum
                        maxValue = .getPersonalKosten(qualifier).Sum
                    End With

                    For s = 1 To anzSnapshots - 1
                        With projekthistorie.ElementAt(s)
                            tmpValue = .getPersonalKosten(qualifier).Sum
                            If tmpValue < minValue Then
                                minIndex = s
                                minValue = tmpValue
                            End If
                            If tmpValue > maxValue Then
                                maxIndex = s
                                maxValue = tmpValue
                            End If
                        End With
                    Next


                Case 5
                    ' Kostenart mit Qualifier
                    With projekthistorie.ElementAt(0)
                        minValue = .getKostenBedarf(qualifier).Sum
                        maxValue = .getKostenBedarf(qualifier).Sum
                    End With

                    For s = 1 To anzSnapshots - 1
                        With projekthistorie.ElementAt(s)
                            tmpValue = .getKostenBedarf(qualifier).Sum
                            If tmpValue < minValue Then
                                minIndex = s
                                minValue = tmpValue
                            End If
                            If tmpValue > maxValue Then
                                maxIndex = s
                                maxValue = tmpValue
                            End If
                        End With
                    Next

                Case Else
                    ' Gesamtkosten
                    With projekthistorie.ElementAt(0)
                        minValue = .getGesamtKostenBedarf.Sum
                        maxValue = .getGesamtKostenBedarf.Sum
                    End With

                    For s = 1 To anzSnapshots - 1
                        With projekthistorie.ElementAt(s)
                            tmpValue = .getGesamtKostenBedarf.Sum
                            If tmpValue < minValue Then
                                minIndex = s
                                minValue = tmpValue
                            End If
                            If tmpValue > maxValue Then
                                maxIndex = s
                                maxValue = tmpValue
                            End If
                        End With
                    Next
            End Select

            Try
                beauftragung = projekthistorie.ElementAt(minIndex)
                lastPlan = projekthistorie.ElementAt(maxIndex)
            Catch ex As Exception
                Throw New ArgumentException("Fehler in Min-/Max Bestimmung " & ex.Message)
            End Try

        End If






        Dim formerEE As Boolean = appInstance.EnableEvents
        appInstance.EnableEvents = False


        Dim minColumn As Integer, maxColumn As Integer, gesternColumn As Integer = getColumnOfDate(heute) - 1
        Dim pastAndFuture As Boolean = False

        Dim werteB(beauftragung.anzahlRasterElemente - 1) As Double
        Dim werteL(lastPlan.anzahlRasterElemente - 1) As Double
        Dim werteC(hproj.anzahlRasterElemente - 1) As Double

        Dim Xdatenreihe() As String
        Dim tdatenreiheB() As Double
        Dim tdatenreiheL() As Double
        Dim tdatenreiheC() As Double

        Dim gesterndatenreihe() As Double
        Dim Xgestern() As String

        ' Bestimmen der Werte 
        Select Case auswahl
            Case 1
                ' Personalkosten
                If isMinMax Then
                    titelTeile(0) = "Min/Max Personalkosten (T€)" & vbLf
                Else
                    titelTeile(0) = "Soll/Ist Personalkosten (T€)" & vbLf
                End If

                kennung = "Soll/Ist Kurve Personalkosten"
                werteB = beauftragung.getAllPersonalKosten
                werteL = lastPlan.getAllPersonalKosten
                werteC = hproj.getAllPersonalKosten
            Case 2
                ' Sonstige Kosten
                If isMinMax Then
                    titelTeile(0) = "Min/Max Sonstige Kosten (T€)" & vbLf
                Else
                    titelTeile(0) = "Soll/Ist Sonstige Kosten (T€)" & vbLf
                End If

                kennung = "Soll/Ist Kurve Sonstige Kosten"
                werteB = beauftragung.getGesamtAndereKosten
                werteL = lastPlan.getGesamtAndereKosten
                werteC = hproj.getGesamtAndereKosten

            Case 3
                ' Gesamt Kosten
                If isMinMax Then
                    titelTeile(0) = "Min/Max Gesamtkosten (T€)" & vbLf
                Else
                    titelTeile(0) = "Soll/Ist Gesamtkosten (T€)" & vbLf
                End If

                kennung = "Soll/Ist Kurve Gesamtkosten"
                werteB = beauftragung.getGesamtKostenBedarf
                werteL = lastPlan.getGesamtKostenBedarf
                werteC = hproj.getGesamtKostenBedarf

            Case 4
                ' Rollen mit Qualifier
                titelTeile(0) = qualifier & "(" & awinSettings.kapaEinheit & ")" & vbLf
                kennung = "Rolle " & qualifier
                Try
                    werteB = beauftragung.getPersonalKosten(qualifier)
                    werteL = lastPlan.getPersonalKosten(qualifier)
                    werteC = hproj.getPersonalKosten(qualifier)
                Catch ex As Exception
                    Throw New ArgumentException(ex.Message & vbLf & qualifier & " nicht gefunden")
                End Try

            Case 5
                ' Kostenart mit Qualifier
                titelTeile(0) = qualifier & " (T€)" & vbLf
                kennung = "Kostenart " & qualifier
                Try
                    werteB = beauftragung.getKostenBedarf(qualifier)
                    werteL = lastPlan.getKostenBedarf(qualifier)
                    werteC = hproj.getKostenBedarf(qualifier)
                Catch ex As Exception
                    Throw New ArgumentException(ex.Message & vbLf & qualifier & " nicht gefunden")
                End Try

            Case Else
                ' Gesamt Kosten
                If isMinMax Then
                    titelTeile(0) = "Min/Max Gesamtkosten (T€)" & vbLf
                Else
                    titelTeile(0) = "Soll/Ist Gesamtkosten (T€)" & vbLf
                End If

                kennung = "Soll/Ist Kurve Gesamtkosten"
                werteB = beauftragung.getGesamtKostenBedarf
                werteL = lastPlan.getGesamtKostenBedarf
                werteC = hproj.getGesamtKostenBedarf
                auswahl = 3

        End Select

        titelTeilLaengen(0) = titelTeile(0).Length
        titelTeile(1) = fullname & vbLf
        titelTeilLaengen(1) = titelTeile(1).Length
        titelTeile(2) = " (" & hproj.timeStamp.ToString & ") "
        titelTeilLaengen(2) = titelTeile(2).Length
        diagramTitle = titelTeile(0) & titelTeile(1) & titelTeile(2)


        minColumn = 10000
        If beauftragung.Start < minColumn Then
            minColumn = beauftragung.Start
        End If

        If lastPlan.Start < minColumn Then
            minColumn = lastPlan.Start
        End If

        If hproj.Start < minColumn Then
            minColumn = hproj.Start
        End If


        With hproj
            maxColumn = .Start + .anzahlRasterElemente - 1
        End With

        With beauftragung
            If maxColumn < .Start + .anzahlRasterElemente - 1 Then
                maxColumn = .Start + .anzahlRasterElemente - 1
            End If
        End With

        With lastPlan
            If maxColumn < .Start + .anzahlRasterElemente - 1 Then
                maxColumn = .Start + .anzahlRasterElemente - 1
            End If
        End With

        ReDim Xdatenreihe(maxColumn - minColumn)
        ReDim tdatenreiheB(maxColumn - minColumn)
        ReDim tdatenreiheL(maxColumn - minColumn)
        ReDim tdatenreiheC(maxColumn - minColumn)


        sumB = 0.0
        sumL = 0.0
        sumC = 0.0
        For i = minColumn To maxColumn
            Xdatenreihe(i - minColumn) = StartofCalendar.AddMonths(i - 1).ToString("MMM yy")
            With beauftragung
                If i >= .Start And i <= .Start + .anzahlRasterElemente - 1 Then
                    tdatenreiheB(i - minColumn) = sumB + werteB(i - .Start)
                    sumB = tdatenreiheB(i - minColumn)
                Else
                    tdatenreiheB(i - minColumn) = sumB
                End If
            End With

            With lastPlan
                If i >= .Start And i <= .Start + .anzahlRasterElemente - 1 Then
                    tdatenreiheL(i - minColumn) = sumL + werteL(i - .Start)
                    sumL = tdatenreiheL(i - minColumn)
                Else
                    tdatenreiheL(i - minColumn) = sumL
                End If
            End With

            With hproj
                If i >= .Start And i <= .Start + .anzahlRasterElemente - 1 Then
                    tdatenreiheC(i - minColumn) = sumC + werteC(i - .Start)
                    sumC = tdatenreiheC(i - minColumn)
                Else
                    tdatenreiheC(i - minColumn) = sumC
                End If
            End With

        Next i


        If gesternColumn >= minColumn And _
            gesternColumn <= maxColumn Then

            pastAndFuture = True
            ReDim gesterndatenreihe(gesternColumn - minColumn)
            ReDim Xgestern(gesternColumn - minColumn)
            For i = minColumn To gesternColumn
                gesterndatenreihe(i - minColumn) = tdatenreiheC(i - minColumn)
                Xgestern(i - minColumn) = Xdatenreihe(i - minColumn)
            Next
        Else
            ReDim gesterndatenreihe(0)
            ReDim Xgestern(0)
            pastAndFuture = False


        End If

        ' jetzt wird das Diagramm gezeichnet 

        With CType(appInstance.Worksheets(arrWsNames(3)), Excel.Worksheet)
            anzDiagrams = CType(.ChartObjects, Excel.ChartObjects).Count
            '
            ' um welches Diagramm handelt es sich ...
            '
            i = 1
            found = False
            While i <= anzDiagrams And Not found
                Dim chtTitle As String
                Try
                    chtTitle = CType(.ChartObjects(i), Excel.ChartObject).Chart.ChartTitle.Text
                Catch ex As Exception
                    chtTitle = " "
                End Try

                If chtTitle = diagramTitle Then
                    found = True

                Else
                    i = i + 1
                End If

            End While

            If found Then
                'Call MsgBox("Chart wird bereits angezeigt ...")
                reportObj = CType(.ChartObjects(i), Excel.ChartObject)
                appInstance.EnableEvents = formerEE
                'appInstance.ScreenUpdating = formerSU
                Exit Sub
            Else
                With appInstance.Charts.Add
                    ' remove extra series
                    Do Until .SeriesCollection.Count = 0
                        .SeriesCollection(1).Delete()
                    Loop

                    .HasAxis(Excel.XlAxisType.xlCategory) = True
                    .HasAxis(Excel.XlAxisType.xlValue) = True

                    'Dim ax As Excel.Axis

                    'With ax
                    '    .MajorUnit = 10000
                    'End With

                    With .Axes(Excel.XlAxisType.xlCategory)
                        .HasTitle = False
                        '.MinimumScale = 0
                        'With .AxisTitle
                        '    .Characters.text = "Monate"
                        '    .Font.Size = 8
                        'End With
                    End With

                    With .Axes(Excel.XlAxisType.xlValue)
                        .HasTitle = False
                        .HasMajorGridlines = False
                        .HasMinorGridlines = False
                    End With

                    .HasLegend = True
                    With .Legend
                        .Position = Excel.Constants.xlTop
                        .Font.Size = awinSettings.fontsizeLegend
                    End With
                    .HasTitle = True
                    .ChartTitle.Text = diagramTitle
                    .ChartTitle.Font.Size = awinSettings.fontsizeTitle
                    .Location(Where:=XlChartLocation.xlLocationAsObject, Name:=appInstance.Worksheets(arrWsNames(3)).name)

                End With

                chtobj = CType(.ChartObjects(anzDiagrams + 1), Excel.ChartObject)
                chtobj.Name = pname & "#" & kennung & "#" & "1"


            End If

            With chtobj.Chart

                .ChartTitle.Format.TextFrame2.TextRange.Characters(titelTeilLaengen(0) + _
                                                                   titelTeilLaengen(1) + 1, titelTeilLaengen(2)).Font.Size = awinSettings.fontsizeLegend

                'series

                If pastAndFuture Then
                    ' dann muss jetzt die "Ist-Markierung gezeichnet werden 

                    With .SeriesCollection.NewSeries
                        .name = "Istwerte"
                        .Interior.color = awinSettings.SollIstFarbeArea
                        .Values = gesterndatenreihe
                        '.XValues = Xgestern
                        .XValues = Xdatenreihe
                        .ChartType = Excel.XlChartType.xlArea

                    End With

                End If

                If isMinMax Or vglBaseline Then
                    With .SeriesCollection.NewSeries
                        If isMinMax Then
                            .name = "Minimum (" & beauftragung.timeStamp.ToString("d") & ")"
                        Else
                            .name = "Soll (" & beauftragung.timeStamp.ToString("d") & ")"
                        End If

                        .Interior.color = awinSettings.SollIstFarbeB
                        .Values = tdatenreiheB
                        .XValues = Xdatenreihe
                        .ChartType = Excel.XlChartType.xlLine
                        .Format.Line.Weight = 2

                    End With
                End If


                If isMinMax Or Not vglBaseline Then
                    With .SeriesCollection.NewSeries
                        If isMinMax Then
                            .name = "Maximum (" & lastPlan.timeStamp.ToString("d") & ")"
                        Else
                            .name = "Last (" & lastPlan.timeStamp.ToString("d") & ")"
                        End If

                        .Interior.color = awinSettings.SollIstFarbeL
                        .Values = tdatenreiheL
                        .XValues = Xdatenreihe
                        .ChartType = Excel.XlChartType.xlLine
                        .Format.Line.Weight = 2

                    End With
                End If


                With .SeriesCollection.NewSeries
                    .name = "Ist (" & hproj.timeStamp.ToString("d") & ")"
                    .Interior.color = awinSettings.SollIstFarbeC
                    .Values = tdatenreiheC
                    .XValues = Xdatenreihe
                    .ChartType = Excel.XlChartType.xlLine
                    .Format.Line.Weight = 2


                End With

            End With


        End With

        With chtobj
            .Top = top
            .Left = left
            .Height = height
            .Width = width
        End With

        appInstance.EnableEvents = formerEE
        reportObj = chtobj

    End Sub

    ''' <summary>
    ''' Methode zeigt zum ausgewählten Projekt die Trendanalyse zu den in der myCollection übergebenen Meilensteinen an 
    ''' </summary>
    ''' <param name="hproj">Verweis auf Projekt</param>
    ''' <param name="repObj">Verweis auf generiertes ChartObject (für Reporting benötigt)</param>
    ''' <param name="myCollection">enthält die NAmen der Meilensteine, für die die Trendanalyse erstellt werden soll</param>
    ''' <param name="top">y-Koordinate linke obere Ecke </param>
    ''' <param name="left">x-Koordinate linke obere Ecke</param>
    ''' <param name="height">Höhe des Charts</param>
    ''' <param name="width">Breite des Charts</param>
    ''' <remarks></remarks>
    Public Sub createMsTrendAnalysisOfProject(ByRef hproj As clsProjekt, ByRef repObj As Excel.ChartObject, ByRef myCollection As Collection, _
                                                 ByVal top As Double, left As Double, height As Double, width As Double)

        Dim kennung As String = " "
        Dim diagramTitle As String = " "
        Dim anzDiagrams As Integer
        Dim found As Boolean
        Dim plen As Integer
        Dim i As Integer
        Dim Xdatenreihe() As String
        Dim tdatenreihe() As Double
        Dim milestoneReached() As Boolean
        Dim prevValueTaken() As Boolean
        Dim ampelfarben() As Long
        Dim tmpdatenreihe() As Date
        Dim chtTitle As String
        Dim pkIndex As Integer = CostDefinitions.Count
        Dim chtobj As Excel.ChartObject
        Dim ErgebnisListeR As New Collection
        Dim msName As String
        Dim zE As String = "(" & awinSettings.kapaEinheit & ")"
        Dim titelTeile(1) As String
        Dim titelTeilLaengen(1) As Integer
        Dim anzMilestones As Integer
        Dim aufzeichnungsStart As Date
        Dim earliestStart As Date
        Dim von As Integer, bis As Integer


        Dim formerEE As Boolean = appInstance.EnableEvents
        Dim formerSU As Boolean = appInstance.ScreenUpdating
        appInstance.EnableEvents = False



        Dim pname As String = hproj.name

        titelTeile(0) = "Meilenstein Trend-Analyse " & hproj.getShapeText & vbLf
        titelTeilLaengen(0) = titelTeile(0).Length
        titelTeile(1) = " (" & hproj.timeStamp.ToString & ") "
        titelTeilLaengen(1) = titelTeile(1).Length
        diagramTitle = titelTeile(0) & titelTeile(1)
        kennung = "MTA"


        ' neu - neu - neu - neu 

        ' wann wurde mit der Aufzeichnung der Projekt-Historie begonnen ? 
        aufzeichnungsStart = projekthistorie.ElementAt(0).timeStamp

        '
        ' bestimme den seit Beauftragung frühesten Start-Monat 
        '
        Try
            earliestStart = projekthistorie.beauftragung.startDate
        Catch ex As Exception
            ' wenn es noch keine Beauftragung gibt, wird das erste Element der Liste verwendet 
            earliestStart = projekthistorie.ElementAt(0).timeStamp
            projekthistorie.currentIndex = 0
        End Try


        ' es beginnt entweder mit dem Monat, wo die Aufzeichnung begann oder mit dem Projekt-Start : nimm das größere von beidem 
        von = System.Math.Max(getColumnOfDate(aufzeichnungsStart), getColumnOfDate(earliestStart))

        ' es endet entweder mit heute oder dem Ende des Projektes : nimm das kleinere von beidem 
        With hproj
            bis = System.Math.Min(getColumnOfDate(Date.Now), getColumnOfDate(.endeDate))
        End With


        ' bestimme die Dimension
        plen = bis - von + 1

        ' wenn nicht mindestens zwei Elemente darstellbar sind, ist kein Trend darzustellen 
        If plen < 2 Then
            appInstance.EnableEvents = formerEE
            Throw New Exception("Es gibt noch keinen Trend für das Projekt '" & hproj.name & "'")
        End If

        ' neu - neu - neu - neu 


        anzMilestones = myCollection.Count

        If anzMilestones = 0 Then
            appInstance.EnableEvents = formerEE
            Throw New Exception("keine Meilensteine angegeben!")
        End If


        ReDim Xdatenreihe(plen - 1)
        ReDim tdatenreihe(plen - 1)
        ReDim ampelfarben(plen - 1)
        ReDim milestoneReached(plen - 1)
        ReDim prevValueTaken(plen - 1)


        For i = 1 To plen
            Xdatenreihe(i - 1) = StartofCalendar.AddMonths(von + i - 2).ToString("MMM yy")
        Next i


        With CType(appInstance.Worksheets(arrWsNames(3)), Excel.Worksheet)
            anzDiagrams = CType(.ChartObjects, Excel.ChartObjects).Count
            '
            ' um welches Diagramm handelt es sich ...
            '
            i = 1
            found = False
            While i <= anzDiagrams And Not found
                Try
                    chtTitle = CType(.ChartObjects(i), Excel.ChartObject).Chart.ChartTitle.Text
                Catch ex As Exception
                    chtTitle = " "
                End Try

                If chtTitle = diagramTitle Then
                    found = True

                Else
                    i = i + 1
                End If

            End While

            If found Then

                appInstance.EnableEvents = formerEE

                repObj = CType(.ChartObjects(i), Excel.ChartObject)
                Exit Sub
            Else
                appInstance.ScreenUpdating = False

                With appInstance.Charts.Add
                    ' remove extra series
                    Do Until .SeriesCollection.Count = 0
                        .SeriesCollection(1).Delete()
                    Loop

                    .HasTitle = True
                    .ChartTitle.Text = diagramTitle
                    .ChartTitle.Font.Size = awinSettings.fontsizeTitle
                    .Location(Where:=XlChartLocation.xlLocationAsObject, Name:=appInstance.Worksheets(arrWsNames(3)).name)

                End With

                chtobj = CType(.ChartObjects(anzDiagrams + 1), Excel.ChartObject)
                chtobj.Name = pname & "#" & kennung & "#" & "1"


            End If

            ' jetzt kommt die Korrektur der Größe; herausfinden, wieviel Raum die Axis Beschriftung einnimmt ... 
            With chtobj
                .Top = top
                .Height = height
                .Left = left
                .Width = width
            End With

            With chtobj.Chart
                ' jetzt wird die Plot-Area so verkleinert, daß links und rechts ausreichend Platz 
                ' für die Bennenung der Meilensteine ist 
                .PlotArea.Left = 0.2 * width
                .PlotArea.Width = 0.8 * width
                '.PlotArea.Height = 0.9 * height
                '.PlotArea.Top = 0.08 * height
            End With

            Dim ms As Integer
            With CType(chtobj.Chart, Excel.Chart)

                .ChartTitle.Format.TextFrame2.TextRange.Characters(titelTeilLaengen(0) + 1, _
                    titelTeilLaengen(1)).Font.Size = awinSettings.fontsizeLegend

                ' remove extra series
                Do Until .SeriesCollection.Count = 0
                    .SeriesCollection(1).Delete()
                Loop

                Dim colorIndex As Integer
                Dim drawnDates As New SortedList(Of Date, Date)
                Dim drawnMilestones As Integer = 0
                Dim tmpMinScale As Date, tmpMaxScale As Date


                For ms = 1 To anzMilestones
                    msName = CStr(myCollection.Item(ms))

                    Try
                        tmpdatenreihe = projekthistorie.getMtaDates(msName, von, bis)
                        drawnMilestones = drawnMilestones + 1

                        ' tmpMinScale und tmpMaxScale werden zur Bestimmung des optimalen Skalierungsfaktors für das Diagramm benötigt 
                        If ms = 1 Then
                            tmpMinScale = tmpdatenreihe.Min
                            tmpMaxScale = tmpdatenreihe.Max
                        Else
                            If tmpMinScale > tmpdatenreihe.Min Then
                                tmpMinScale = tmpdatenreihe.Min
                            End If

                            If tmpMaxScale < tmpdatenreihe.Max Then
                                tmpMaxScale = tmpdatenreihe.Max
                            End If
                        End If

                        ReDim tdatenreihe(tmpdatenreihe.Length - 1)
                        For qx = 0 To tmpdatenreihe.Length - 1
                            ' nur der Datums-Wert ohne Zeit-Anteil - die Farbe ist als Anzahl Sekunden nach Tagesstart in das Datum kodiert ...
                            tdatenreihe(qx) = tmpdatenreihe(qx).Date.ToOADate

                            ' prüfen, ob der Wert vonm Vormonat übernommen wurde 
                            If DateDiff(DateInterval.Hour, tmpdatenreihe(qx).Date, tmpdatenreihe(qx)) > 11 Then
                                prevValueTaken(qx) = True
                                tmpdatenreihe(qx) = tmpdatenreihe(qx).AddHours(-12) ' Kodierung für "Wert des Vormonats" rausnehmen, sonst ist nachher Farbe auf alle Fälle rot
                            Else
                                prevValueTaken(qx) = False
                            End If

                            If DateDiff(DateInterval.Hour, tmpdatenreihe(qx).Date, tmpdatenreihe(qx)) = 6 Then
                                milestoneReached(qx) = True
                                tmpdatenreihe(qx) = tmpdatenreihe(qx).AddHours(-6) ' Kodierung für "milestone abgeschlossen" rausnehmen, sonst ist nachher Farbe auf alle Fälle rot
                            Else
                                milestoneReached(qx) = False
                            End If

                            colorIndex = CInt(DateDiff(DateInterval.Second, tmpdatenreihe(qx).Date, tmpdatenreihe(qx)))
                            If colorIndex = 0 Then
                                ampelfarben(qx) = awinSettings.AmpelNichtBewertet
                            ElseIf colorIndex = 1 Then
                                ampelfarben(qx) = awinSettings.AmpelGruen
                            ElseIf colorIndex = 2 Then
                                ampelfarben(qx) = awinSettings.AmpelGelb
                            Else
                                ampelfarben(qx) = awinSettings.AmpelRot
                            End If
                        Next

                        'series
                        With CType(.SeriesCollection.NewSeries, Excel.Series)
                            .Name = drawnMilestones.ToString & " - " & elemNameOfElemID(msName)
                            .ChartType = Excel.XlChartType.xlLineMarkers
                            .Interior.Color = awinSettings.AmpelNichtBewertet
                            .Values = tdatenreihe
                            .XValues = Xdatenreihe
                            .HasDataLabels = False
                            .MarkerStyle = Excel.XlMarkerStyle.xlMarkerStyleCircle
                            .MarkerForegroundColor = CInt(awinSettings.AmpelNichtBewertet)
                            .MarkerBackgroundColor = CInt(awinSettings.AmpelNichtBewertet)

                            With .Format.Line
                                .Visible = MsoTriState.msoTrue
                                .ForeColor.RGB = CInt(awinSettings.AmpelNichtBewertet)
                                .DashStyle = MsoLineDashStyle.msoLineDashDot
                            End With
                        End With


                        For px = 1 To tdatenreihe.Length

                            With CType(.SeriesCollection(drawnMilestones).Points(px), Point)
                                .Interior.Color = ampelfarben(px - 1)
                                .MarkerStyle = Excel.XlMarkerStyle.xlMarkerStyleCircle
                                .MarkerForegroundColor = CInt(ampelfarben(px - 1))
                                .MarkerBackgroundColor = CInt(ampelfarben(px - 1))
                                .MarkerSize = 10

                                ' Schreiben des ersten Planungs-Standes
                                If px = 1 Then

                                    ' wenn es der Wert aus dem Vormonat ist: einen kleineren Marker zeichnen 
                                    If prevValueTaken(px - 1) Then
                                        .Interior.Color = CInt(awinSettings.AmpelNichtBewertet)
                                        .MarkerForegroundColor = CInt(awinSettings.AmpelNichtBewertet)
                                        .MarkerBackgroundColor = CInt(awinSettings.AmpelNichtBewertet)
                                        .MarkerSize = 2
                                    End If

                                    ' wenn der Meilenstein zum zeitpunkt des Planungs-Standes bereits in der Vergangenheit lag, wird er auch so markiert
                                    If milestoneReached(px - 1) Then
                                        .MarkerStyle = Excel.XlMarkerStyle.xlMarkerStyleSquare
                                    End If

                                    .HasDataLabel = True
                                    If anzMilestones > 1 Then
                                        .DataLabel.Text = drawnMilestones.ToString & " - " & tmpdatenreihe(px - 1).ToShortDateString
                                    Else
                                        .DataLabel.Text = elemNameOfElemID(msName) & vbLf & tmpdatenreihe(px - 1).ToShortDateString
                                    End If

                                    .DataLabel.Font.Size = awinSettings.fontsizeItems
                                    Try
                                        .DataLabel.Position = Excel.XlDataLabelPosition.xlLabelPositionLeft
                                    Catch ex As Exception

                                    End Try

                                    Try
                                        drawnDates.Add(tmpdatenreihe(px - 1).Date, tmpdatenreihe(px - 1))
                                    Catch ex As Exception

                                    End Try
                                End If


                                ' wenn mittendrin Daten auftauchen, die noch nicht geschrieben wurden ... 
                                If px > 1 And px < tdatenreihe.Length Then

                                    ' wenn es der Wert aus dem Vormonat ist: einen kleineren Marker zeichnen 
                                    If prevValueTaken(px - 1) Then
                                        .Interior.Color = CInt(awinSettings.AmpelNichtBewertet)
                                        .MarkerForegroundColor = CInt(awinSettings.AmpelNichtBewertet)
                                        .MarkerBackgroundColor = CInt(awinSettings.AmpelNichtBewertet)
                                        .MarkerSize = 2
                                    End If

                                    ' wenn der Meilenstein zum zeitpunkt des Planungs-Standes bereits in der Vergangenheit lag, wird er entsprechend markiert
                                    If milestoneReached(px - 1) Then
                                        .MarkerStyle = Excel.XlMarkerStyle.xlMarkerStyleSquare
                                    End If

                                    If Not drawnDates.ContainsKey(tmpdatenreihe(px - 1).Date) And _
                                        tmpdatenreihe(px - 1).Date <> tmpdatenreihe(tdatenreihe.Length - 1).Date Then

                                        .HasDataLabel = True
                                        .DataLabel.Text = tmpdatenreihe(px - 1).ToShortDateString
                                        '.DataLabel.Text = msName & ": " & tmpdatenreihe(px - 1).ToShortDateString
                                        .DataLabel.Font.Size = awinSettings.fontsizeItems
                                        Try

                                            If tmpdatenreihe(px - 1).Date > tmpdatenreihe(px - 2).Date Then
                                                .DataLabel.Position = Excel.XlDataLabelPosition.xlLabelPositionBelow
                                            Else
                                                .DataLabel.Position = Excel.XlDataLabelPosition.xlLabelPositionAbove
                                            End If

                                        Catch ex As Exception
                                            .DataLabel.Position = Excel.XlDataLabelPosition.xlLabelPositionCenter
                                        End Try

                                        Try
                                            drawnDates.Add(tmpdatenreihe(px - 1).Date, tmpdatenreihe(px - 1))
                                        Catch ex As Exception

                                        End Try

                                    End If
                                End If

                                ' Schreiben des letzten Planungs-Standes
                                If px > 1 And px = tdatenreihe.Length Then

                                    ' wenn es der Wert aus dem Vormonat ist: einen kleineren/ nicht sichtbaren Marker zeichnen 
                                    If prevValueTaken(px - 1) Then
                                        .Interior.Color = CInt(awinSettings.AmpelNichtBewertet)
                                        .MarkerForegroundColor = CInt(awinSettings.AmpelNichtBewertet)
                                        .MarkerBackgroundColor = CInt(awinSettings.AmpelNichtBewertet)
                                        .MarkerSize = 2
                                    End If

                                    ' wenn der Meilenstein zum zeitpunkt des Planungs-Standes bereits in der Vergangenheit lag, wird er auch so markiert
                                    If milestoneReached(px - 1) Then
                                        .MarkerStyle = Excel.XlMarkerStyle.xlMarkerStyleSquare
                                    End If

                                    .HasDataLabel = True
                                    If anzMilestones > 1 Then
                                        .DataLabel.Text = drawnMilestones.ToString & " - " & tmpdatenreihe(px - 1).ToShortDateString
                                    Else
                                        .DataLabel.Text = elemNameOfElemID(msName) & vbLf & tmpdatenreihe(px - 1).ToShortDateString
                                    End If
                                    .DataLabel.Font.Size = awinSettings.fontsizeItems
                                    Try
                                        .DataLabel.Position = Excel.XlDataLabelPosition.xlLabelPositionRight
                                    Catch ex As Exception

                                    End Try

                                    Try
                                        drawnDates.Add(tmpdatenreihe(px - 1).Date, tmpdatenreihe(px - 1))
                                    Catch ex As Exception

                                    End Try


                                End If

                            End With
                        Next

                        drawnDates.Clear()


                    Catch ex As Exception

                    End Try




                Next ms

                ' Bestimmen des optimlaen skalierungsfaktors 
                Dim spread As Integer
                spread = CInt(DateDiff(DateInterval.Day, tmpMinScale, tmpMaxScale) / 10)
                If spread < 1 Then
                    spread = 1
                End If

                tmpMinScale = tmpMinScale.AddDays(-1 * spread)
                tmpMaxScale = tmpMaxScale.AddDays(spread)

                .HasAxis(Excel.XlAxisType.xlCategory) = True
                .HasAxis(Excel.XlAxisType.xlValue) = False


                With CType(.Axes(Excel.XlAxisType.xlCategory), Excel.Axis)
                    .HasTitle = True
                    .AxisTitle.Text = "Berichtszeiträume"
                    .AxisTitle.Format.TextFrame2.TextRange.Font.Size = 14
                    .BaseUnit = Excel.XlTimeUnit.xlMonths
                    .CategoryType = Excel.XlCategoryType.xlTimeScale
                End With

                With CType(.Axes(Excel.XlAxisType.xlValue), Excel.Axis)
                    .HasMajorGridlines = False

                    Try
                        .HasTitle = False
                        '.AxisTitle.Text = "Meilenstein Termine"
                        '.AxisTitle.Format.TextFrame2.TextRange.Font.Size = 14
                    Catch ex As Exception

                    End Try

                    .MaximumScale = tmpMaxScale.ToOADate
                    .MinimumScale = tmpMinScale.ToOADate
                    .MajorUnit = 61

                    'Try
                    '    .TickLabels.NumberFormat = "dd-mm-yyyy"
                    'Catch ex As Exception

                    'End Try


                End With

                If anzMilestones > 1 Then
                    .HasLegend = True
                    With .Legend
                        .Position = Excel.XlLegendPosition.xlLegendPositionTop
                        .Font.Size = awinSettings.fontsizeLegend
                    End With
                Else
                    .HasLegend = False
                End If



            End With



        End With



        'Call awinScrollintoView()
        appInstance.EnableEvents = formerEE
        appInstance.ScreenUpdating = formerSU

        repObj = chtobj



    End Sub

    

    ''' <summary>
    ''' Prozedur zeigt die Ressourcen Struktur des Projektes an (Balken-Diagramm)
    ''' </summary>
    ''' <param name="hproj"></param>
    ''' <param name="repObj">Verweis auf das Grafik Objekt. 
    ''' Das wird dann von der Reporting Engine verwendet </param>
    ''' <param name="auswahl">steuert, was angezeigt wird
    ''' Auswahl = 1 : Diagramm zeigt Mann-Monate
    ''' Auswahl = 2 : Diagramm zeigt Personal-Kosten
    ''' </param>
    ''' <param name="top"></param>
    ''' <param name="left"></param>
    ''' <param name="height"></param>
    ''' <param name="width"></param>
    ''' <remarks>Kennung Phasen, Personalbedarf, Personalkosten, Sonstige Kosten, Gesamtkosten, Strategie, Ergebnis</remarks>
    Public Sub createRessBalkenOfProject(ByRef hproj As clsProjekt, ByRef repObj As Excel.ChartObject, ByVal auswahl As Integer, _
                                            ByVal top As Double, left As Double, height As Double, width As Double)

        Dim kennung As String = " "
        Dim diagramTitle As String = " "
        Dim anzDiagrams As Integer
        Dim found As Boolean
        Dim plen As Integer
        Dim i As Integer
        Dim Xdatenreihe() As String
        Dim tdatenreihe() As Double
        Dim hsum() As Double, gesamt_summe As Double
        Dim anzRollen As Integer
        Dim chtTitle As String
        Dim pkIndex As Integer = CostDefinitions.Count
        Dim pstart As Integer
        Dim chtobj As Excel.ChartObject
        Dim ErgebnisListeR As New Collection
        Dim roleName As String
        Dim zE As String = "(" & awinSettings.kapaEinheit & ")"
        Dim titelTeile(1) As String
        Dim titelTeilLaengen(1) As Integer
        Dim tmpcollection As New Collection

        Dim formerEE As Boolean = appInstance.EnableEvents
        'Dim formerSU As Boolean = appInstance.ScreenUpdating
        appInstance.EnableEvents = False
        'appInstance.ScreenUpdating = False


        Dim pname As String = hproj.name

        tmpcollection.Add(hproj.getShapeText & "#" & auswahl.ToString)
        kennung = calcChartKennung("pr", PTprdk.PersonalBalken, tmpcollection)

        If auswahl = 1 Then
            titelTeile(0) = "Personalbedarf " & zE & vbLf & hproj.getShapeText & vbLf
            titelTeilLaengen(0) = titelTeile(0).Length
            titelTeile(1) = " (" & hproj.timeStamp.ToString & ") "
            titelTeilLaengen(1) = titelTeile(1).Length
            diagramTitle = titelTeile(0) & titelTeile(1)
            'kennung = "Personalbedarf"
        ElseIf auswahl = 2 Then
            titelTeile(0) = "Personalkosten (T€)" & vbLf & hproj.getShapeText & vbLf
            titelTeilLaengen(0) = titelTeile(0).Length
            titelTeile(1) = " (" & hproj.timeStamp.ToString & ") "
            titelTeilLaengen(1) = titelTeile(1).Length
            diagramTitle = titelTeile(0) & titelTeile(1)
            diagramTitle = titelTeile(0) & titelTeile(1)
            'kennung = "Personalkosten"
        Else
            diagramTitle = "--- (T€)" & vbLf & pname
            'kennung = "Gesamtkosten"
        End If



        '
        ' hole die Projektdauer
        '
        With hproj
            plen = .anzahlRasterElemente
            pstart = .Start
        End With
        '
        ' hole die Anzahl Kostenarten, die in diesem Projekt vorkommen
        '
        '
        ' hole die Anzahl Rollen, die in diesem Projekt vorkommen
        '
        ErgebnisListeR = hproj.getUsedRollen
        anzRollen = ErgebnisListeR.Count

        If anzRollen = 0 Then
            Throw New Exception("keine Ressourcen Bedarfe definiert")
        End If


        ReDim Xdatenreihe(plen - 1)
        ReDim tdatenreihe(plen - 1)


        ReDim hsum(anzRollen - 1)


        For i = 1 To plen
            Xdatenreihe(i - 1) = StartofCalendar.AddMonths(pstart + i - 2).ToString("MMM yy")
        Next i

        gesamt_summe = 0
        With CType(appInstance.Worksheets(arrWsNames(3)), Excel.Worksheet)
            anzDiagrams = CType(.ChartObjects, Excel.ChartObjects).Count
            '
            ' um welches Diagramm handelt es sich ...
            '
            i = 1
            found = False
            While i <= anzDiagrams And Not found
                Try
                    chtTitle = CType(.ChartObjects(i), Excel.ChartObject).Chart.ChartTitle.Text
                Catch ex As Exception
                    chtTitle = " "
                End Try

                If chtTitle = diagramTitle Then
                    found = True

                Else
                    i = i + 1
                End If

            End While

            If found Then
                'Call MsgBox("Chart wird bereits angezeigt ...")
                appInstance.EnableEvents = formerEE
                'appInstance.ScreenUpdating = formerSU
                repObj = CType(.ChartObjects(i), Excel.ChartObject)
                Exit Sub
            Else
                With appInstance.Charts.Add
                    ' remove extra series
                    Do Until .SeriesCollection.Count = 0
                        .SeriesCollection(1).Delete()
                    Loop

                    .HasAxis(Excel.XlAxisType.xlCategory) = True
                    .HasAxis(Excel.XlAxisType.xlValue) = True

                    With .Axes(Excel.XlAxisType.xlCategory)
                        .HasTitle = False
                        '.MinimumScale = 0
                        'With .AxisTitle
                        '    .Characters.text = "Monate"
                        '    .Font.Size = 8
                        'End With
                    End With

                    With .Axes(Excel.XlAxisType.xlValue)
                        .HasTitle = False
                        '.MaximumScale = maxscale
                        '.MinimumScale = 0

                        'With .AxisTitle
                        '    .Characters.text = "Kosten"
                        '    .Font.Size = 8
                        'End With
                    End With

                    .HasLegend = True
                    With .Legend
                        .Position = Excel.Constants.xlTop
                        .Font.Size = awinSettings.fontsizeLegend
                    End With
                    .HasTitle = True
                    .ChartTitle.Text = diagramTitle
                    .ChartTitle.Font.Size = awinSettings.fontsizeTitle
                    .Location(Where:=XlChartLocation.xlLocationAsObject, Name:=appInstance.Worksheets(arrWsNames(3)).name)

                End With

                chtobj = CType(.ChartObjects(anzDiagrams + 1), Excel.ChartObject)
                'chtobj.Name = pname & "#" & kennung & "#" & "1"
                chtobj.Name = kennung



            End If

            With chtobj.Chart

                .ChartTitle.Format.TextFrame2.TextRange.Characters(titelTeilLaengen(0) + 1, _
                    titelTeilLaengen(1)).Font.Size = awinSettings.fontsizeLegend

                For r = 1 To anzRollen
                    roleName = CStr(ErgebnisListeR.Item(r))
                    If auswahl = 1 Then
                        tdatenreihe = hproj.getRessourcenBedarf(roleName)
                    Else
                        tdatenreihe = hproj.getPersonalKosten(roleName)
                    End If
                    hsum(r - 1) = 0
                    For i = 0 To plen - 1
                        hsum(r - 1) = hsum(r - 1) + tdatenreihe(i)
                    Next i
                    gesamt_summe = gesamt_summe + hsum(r - 1)

                    'series
                    With .SeriesCollection.NewSeries
                        .name = roleName
                        .Interior.color = RoleDefinitions.getRoledef(roleName).farbe
                        .Values = tdatenreihe
                        .XValues = Xdatenreihe
                        .ChartType = Excel.XlChartType.xlColumnStacked
                    End With

                Next r


            End With

            ' jetzt kommt die Korrektur der Größe; herausfinden, wieviel Raum die Axis Beschriftung einnimmt ... 
            With chtobj
                .Top = top
                .Height = 2 * height

                Dim axleft As Double, axwidth As Double
                If .Chart.HasAxis(Excel.XlAxisType.xlValue) = True Then
                    With CType(.Chart.Axes(Excel.XlAxisType.xlValue), Excel.Axis)
                        axleft = .Left
                        axwidth = .Width
                    End With
                    If left - axwidth < 1 Then
                        left = 1
                        width = width + left + 9
                    Else
                        left = left - axwidth
                        width = width + axwidth + 9
                    End If

                End If

                .Left = left
                .Width = width


            End With

        End With



        'Call awinScrollintoView()
        appInstance.EnableEvents = formerEE
        'appInstance.ScreenUpdating = formerSU

        repObj = chtobj


    End Sub


    ''' <summary>
    ''' aktualisiert das Info Chart "Ressourcen Bedarf eines Projektes
    ''' übergeben wird das Projekt sowie das Chart 
    ''' </summary>
    ''' <param name="hproj">selektiertes Projekt</param>
    ''' <param name="chtobj">Chart, das ein Projekt-Info Chart darstellt</param>
    ''' <param name="auswahl">
    ''' 1: Diagramm zeigt Mann-Monate 
    ''' 2: Diagramm zeigt Personal-Kosten</param>
    ''' <param name="changeScale">gibt an, ob der Scale ggf an neue Werte angepasst werden muss</param>
    ''' <remarks>wenn es aus der Time Machine aus aufgerufen wird, darf der Scale gerade nicht angepasst werden </remarks>
    Public Sub updateRessBalkenOfProject(ByRef hproj As clsProjekt, ByRef chtobj As Excel.ChartObject, _
                                         ByVal auswahl As Integer, ByVal changeScale As Boolean)


        Dim kennung As String = " "
        Dim diagramTitle As String = " "
        Dim plen As Integer
        Dim i As Integer
        Dim Xdatenreihe() As String
        Dim tdatenreihe() As Double
        Dim sumdatenreihe() As Double
        'Dim hsum() As Double, gesamt_summe As Double
        Dim anzRollen As Integer
        Dim pkIndex As Integer = CostDefinitions.Count
        Dim pstart As Integer
        Dim ErgebnisListeR As New Collection
        Dim roleName As String
        Dim zE As String = "(" & awinSettings.kapaEinheit & ")"
        Dim titelTeile(1) As String
        Dim titelTeilLaengen(1) As Integer
        Dim tmpCollection As New Collection
        Dim formerEE As Boolean = appInstance.EnableEvents
        appInstance.EnableEvents = False


        Dim pname As String = hproj.name

        tmpCollection.Add(hproj.getShapeText & "#" & auswahl.ToString)
        kennung = calcChartKennung("pr", PTprdk.PersonalBalken, tmpCollection)

        If auswahl = 1 Then
            titelTeile(0) = "Personalbedarf " & zE & vbLf & hproj.getShapeText & vbLf
            titelTeilLaengen(0) = titelTeile(0).Length
            titelTeile(1) = " (" & hproj.timeStamp.ToString & ") "
            titelTeilLaengen(1) = titelTeile(1).Length
            diagramTitle = titelTeile(0) & titelTeile(1)
            'kennung = "Personalbedarf"
        ElseIf auswahl = 2 Then
            titelTeile(0) = "Personalkosten (T€)" & vbLf & hproj.getShapeText & vbLf
            titelTeilLaengen(0) = titelTeile(0).Length
            titelTeile(1) = " (" & hproj.timeStamp.ToString & ") "
            titelTeilLaengen(1) = titelTeile(1).Length
            diagramTitle = titelTeile(0) & titelTeile(1)
            'kennung = "Personalkosten"
        Else
            diagramTitle = "--- (T€)" & vbLf & pname
            'kennung = "Gesamtkosten"
        End If



        '
        ' hole die Projektdauer
        '
        With hproj
            plen = .anzahlRasterElemente
            pstart = .Start
        End With
        '
        ' hole die Anzahl Kostenarten, die in diesem Projekt vorkommen
        '
        '
        ' hole die Anzahl Rollen, die in diesem Projekt vorkommen
        '
        ErgebnisListeR = hproj.getUsedRollen
        anzRollen = ErgebnisListeR.Count




        ReDim Xdatenreihe(plen - 1)
        ReDim tdatenreihe(plen - 1)
        ReDim sumdatenreihe(plen - 1)

        ' sonst kommt der in eine Endlos Schleife, wenn keine Rollen definiert sind 
        'If anzRollen > 0 Then
        '    ReDim hsum(anzRollen - 1)
        'Else
        '    ReDim hsum(0)
        'End If


        For i = 1 To plen
            Xdatenreihe(i - 1) = StartofCalendar.AddMonths(pstart + i - 2).ToString("MMM yy")
        Next i

        'gesamt_summe = 0


        With chtobj.Chart

            '' remove extra series
            Do Until .SeriesCollection.Count = 0
                .SeriesCollection(1).Delete()
            Loop

            For r = 1 To anzRollen
                roleName = CStr(ErgebnisListeR.Item(r))
                If auswahl = 1 Then
                    tdatenreihe = hproj.getRessourcenBedarf(roleName)
                Else
                    tdatenreihe = hproj.getPersonalKosten(roleName)
                End If

                For i = 0 To plen - 1
                    sumdatenreihe(i) = sumdatenreihe(i) + tdatenreihe(i)
                Next
                'hsum(r - 1) = 0
                'For i = 0 To plen - 1
                '    hsum(r - 1) = hsum(r - 1) + tdatenreihe(i)
                'Next i
                'gesamt_summe = gesamt_summe + hsum(r - 1)

                'series
                With .SeriesCollection.NewSeries

                    .Name = roleName
                    .Interior.color = RoleDefinitions.getRoledef(roleName).farbe
                    .Values = tdatenreihe
                    .XValues = Xdatenreihe
                    .ChartType = Excel.XlChartType.xlColumnStacked
                End With

            Next r

            If CBool(.HasAxis(Excel.XlAxisType.xlValue)) Then

                With CType(.Axes(Excel.XlAxisType.xlValue), Excel.Axis)
                    ' das ist dann relevant, wenn ein anderes Projekt selektiert wird, das über die aktuelle Skalierung 
                    ' hinausgehende Werte hat 

                    If changeScale Then
                        .MinimumScale = 0
                        If Not (.MaximumScaleIsAuto) Then
                            Dim tstValue As Double = .MaximumScale

                            If sumdatenreihe.Max > .MaximumScale - 3 Then
                                .MaximumScale = sumdatenreihe.Max + 3
                            End If
                            .MaximumScaleIsAuto = True
                        End If
                    End If

                End With

            End If


            If .HasTitle Then
                .ChartTitle.Text = diagramTitle
                ' ur: 21.07.2014: auskommentiert für Chart-Cockpit
                '.ChartTitle.Font.Size = awinSettings.fontsizeTitle
                '.ChartTitle.Format.TextFrame2.TextRange.Characters(titelTeilLaengen(0) + 1, _
                '       titelTeilLaengen(1)).Font.Size = awinSettings.fontsizeLegend
            End If


        End With

        chtobj.Name = kennung



        appInstance.EnableEvents = formerEE
        'appInstance.ScreenUpdating = formerSU



    End Sub
    '
    ' Prozedur zeigt die Kosten Struktur des Projektes an (Balken-Diagramm)
    '
    ' Auswahl = 1 : Diagramm zeigt nur sonstige Kosten 
    ' Auswahl = 2 : Diagramm zeigt alle Kosten, inkl Personalkosten 
    ' kennziffer = 0 : Phasen Diagramm
    '            = 1 : Personal-Bedarfe (Balken)
    '            = 2 : Personal-Bedarfe (PIE)
    '            = 3 : Kosten (Balken)
    '            = 4 : Kosten (Pie)
    '            = 5 : Strategie / Risiko 
    '            = 6 : Ergebnis

    Public Sub createCostBalkenOfProject(ByRef hproj As clsProjekt, ByRef repObj As Excel.ChartObject, ByVal auswahl As Integer, _
                                        ByVal top As Double, left As Double, height As Double, width As Double)

        Dim kennung As String
        Dim diagramTitle As String
        Dim anzDiagrams As Integer
        Dim found As Boolean
        Dim plen As Integer
        Dim i As Integer
        Dim Xdatenreihe() As String
        Dim tdatenreihe() As Double
        Dim hsum() As Double, gesamt_summe As Double
        Dim anzKostenarten As Integer
        Dim costname As String
        Dim chtTitle As String
        Dim pkIndex As Integer = CostDefinitions.Count
        Dim pstart As Integer
        Dim chtobj As Excel.ChartObject
        Dim titelTeile(1) As String
        Dim titelTeilLaengen(1) As Integer
        Dim tmpcollection As New Collection


        Dim ErgebnisListeK As Collection

        Dim formerEE As Boolean = appInstance.EnableEvents
        appInstance.EnableEvents = False


        Dim pname As String = hproj.name

        tmpcollection.Add(hproj.getShapeText & "#" & auswahl.ToString)
        kennung = calcChartKennung("pr", PTprdk.KostenBalken, tmpcollection)

        If auswahl = 1 Then

            titelTeile(0) = "Sonstige Kosten T€" & vbLf & hproj.getShapeText & vbLf
            titelTeilLaengen(0) = titelTeile(0).Length
            titelTeile(1) = " (" & hproj.timeStamp.ToString & ") "
            titelTeilLaengen(1) = titelTeile(1).Length
            diagramTitle = titelTeile(0) & titelTeile(1)
            'kennung = "Sonstige Kosten"
        Else
            titelTeile(0) = "Gesamtkosten T€" & vbLf & hproj.getShapeText & vbLf
            titelTeilLaengen(0) = titelTeile(0).Length
            titelTeile(1) = " (" & hproj.timeStamp.ToString & ") "
            titelTeilLaengen(1) = titelTeile(1).Length
            diagramTitle = titelTeile(0) & titelTeile(1)
            'kennung = "Gesamtkosten"
        End If


        '
        ' hole die Projektdauer
        '
        With hproj
            plen = .anzahlRasterElemente
            pstart = .Start
        End With
        '
        ' hole die Anzahl Kostenarten, die in diesem Projekt vorkommen
        '
        ErgebnisListeK = hproj.getUsedKosten
        anzKostenarten = ErgebnisListeK.Count

        ' es wird die Null angezeigt 
        'If anzKostenarten = 0 And auswahl = 1 Then
        '    MsgBox("keine Kosten-Bedarfe definiert")
        '    appInstance.EnableEvents = formerEE
        '    'appInstance.ScreenUpdating = formerSU
        '    Exit Sub
        'End If


        ReDim Xdatenreihe(plen - 1)
        ReDim tdatenreihe(plen - 1)


        If auswahl = 1 Then

            If anzKostenarten = 0 Then
                ReDim hsum(0)
            Else
                ReDim hsum(anzKostenarten - 1)
            End If

        Else
            ReDim hsum(anzKostenarten) ' weil jetzt die berechneten Personalkosten dazu kommen
        End If


        For i = 1 To plen
            Xdatenreihe(i - 1) = StartofCalendar.AddMonths(pstart + i - 2).ToString("MMM yy")
        Next i


        Dim ik As Integer = 1 ' wird für die Unterscheidung benötigt, ob mit Personal-Kosten oder ohne 
        gesamt_summe = 0
        With CType(appInstance.Worksheets(arrWsNames(3)), Excel.Worksheet)
            anzDiagrams = CType(.ChartObjects, Excel.ChartObjects).Count
            '
            ' um welches Diagramm handelt es sich ...
            '
            i = 1
            found = False
            While i <= anzDiagrams And Not found
                Try
                    chtTitle = CType(.ChartObjects(i), Excel.ChartObject).Chart.ChartTitle.Text
                Catch ex As Exception
                    chtTitle = " "
                End Try

                If chtTitle = diagramTitle Then
                    found = True

                Else
                    i = i + 1
                End If

            End While

            If found Then
                'Call MsgBox("Chart wird bereits angezeigt ...")
                appInstance.EnableEvents = formerEE
                repObj = CType(.ChartObjects(i), Excel.ChartObject)
                'appInstance.ScreenUpdating = formerSU
                Exit Sub
            Else
                With appInstance.Charts.Add
                    ' remove extra series
                    Do Until .SeriesCollection.Count = 0
                        .SeriesCollection(1).Delete()
                    Loop

                    .HasAxis(Excel.XlAxisType.xlCategory) = True
                    .HasAxis(Excel.XlAxisType.xlValue) = True

                    With .Axes(Excel.XlAxisType.xlCategory)
                        .HasTitle = False
                        '.MinimumScale = 0
                        'With .AxisTitle
                        '    .Characters.text = "Monate"
                        '    .Font.Size = 8
                        'End With
                    End With

                    With .Axes(Excel.XlAxisType.xlValue)
                        .HasTitle = False
                        '.MaximumScale = maxscale
                        '.MinimumScale = 0

                        'With .AxisTitle
                        '    .Characters.text = "Kosten"
                        '    .Font.Size = 8
                        'End With
                    End With

                    .HasLegend = True
                    With .Legend
                        .Position = Excel.Constants.xlTop
                        .Font.Size = awinSettings.fontsizeLegend
                    End With
                    .HasTitle = True
                    .ChartTitle.Text = diagramTitle
                    .ChartTitle.Font.Size = awinSettings.fontsizeTitle
                    .ChartTitle.Format.TextFrame2.TextRange.Characters(titelTeilLaengen(0) + 1, _
                        titelTeilLaengen(1)).Font.Size = awinSettings.fontsizeLegend


                    .Location(Where:=XlChartLocation.xlLocationAsObject, Name:=appInstance.Worksheets(arrWsNames(3)).name)

                End With

                chtobj = CType(.ChartObjects(anzDiagrams + 1), Excel.ChartObject)
                chtobj.Name = kennung



            End If

            With chtobj.Chart


                If auswahl = 2 Then
                    ik = 0
                    costname = "Personal-Kosten"
                    tdatenreihe = hproj.getAllPersonalKosten
                    hsum(ik) = 0
                    For i = 0 To plen - 1
                        hsum(ik) = hsum(ik) + tdatenreihe(i)
                    Next i

                    gesamt_summe = gesamt_summe + hsum(ik)


                    With .SeriesCollection.NewSeries
                        .name = costname
                        .Interior.color = CostDefinitions.getCostdef(pkIndex).farbe
                        .Values = tdatenreihe
                        .XValues = Xdatenreihe
                        .ChartType = Excel.XlChartType.xlColumnStacked
                    End With
                End If

                For k = 1 To anzKostenarten
                    costname = CStr(ErgebnisListeK.Item(k))
                    tdatenreihe = hproj.getKostenBedarf(costname)
                    hsum(k - ik) = 0
                    For i = 0 To plen - 1
                        hsum(k - ik) = hsum(k - ik) + tdatenreihe(i)
                    Next i

                    gesamt_summe = gesamt_summe + hsum(k - ik)

                    With .SeriesCollection.NewSeries
                        .name = costname
                        .Interior.color = CostDefinitions.getCostdef(costname).farbe
                        .Values = tdatenreihe
                        .XValues = Xdatenreihe
                        .ChartType = Excel.XlChartType.xlColumnStacked
                    End With

                Next k


            End With

            ' jetzt kommt die Korrektur der Größe; herausfinden, wieviel Raum die Axis Beschriftung einnimmt ... 
            With chtobj
                .Top = top
                .Height = 2 * height

                Dim axleft As Double, axwidth As Double
                If .Chart.HasAxis(Excel.XlAxisType.xlValue) = True Then
                    With CType(.Chart.Axes(Excel.XlAxisType.xlValue), Excel.Axis)
                        axleft = .Left
                        axwidth = .Width
                    End With
                    If left - axwidth < 1 Then
                        left = 1
                        width = width + left + 9
                    Else
                        left = left - axwidth
                        width = width + axwidth + 9
                    End If

                End If

                .Left = left
                .Width = width


            End With

        End With



        'Call awinScrollintoView()
        appInstance.EnableEvents = formerEE
        'appInstance.ScreenUpdating = formerSU

        repObj = chtobj

    End Sub

    ''' <summary>
    ''' aktualisiert das Auslastungs Chart mit den Über- bzw Unterauslastungs-Details
    ''' </summary>
    ''' <param name="chtobj">Verweis auf Excel Chart Objekt</param>
    ''' <param name="auswahl">
    '''         = 1 : Überauslastung
    '''         = 2 : Unterauslastung
    ''' </param>
    ''' <remarks></remarks>
    Public Sub updateAuslastungsDetailPie(ByRef chtobj As Excel.ChartObject, ByVal auswahl As Integer)

        Dim diagramTitle As String

        Dim Xdatenreihe() As String
        Dim tdatenreihe() As Double

        Dim anzRollen As Integer
        Dim roleName As String


        Dim zE As String = awinSettings.kapaEinheit
        Dim titelTeile(1) As String
        Dim titelTeilLaengen(1) As Integer


        Dim formerEE As Boolean = appInstance.EnableEvents
        appInstance.EnableEvents = False


        '
        ' hole die Anzahl Rollen
        '
        anzRollen = RoleDefinitions.Count

        If anzRollen = 0 Then
            MsgBox("keine Rollen-Bedarfe definiert")
            Exit Sub
        End If

        ReDim tdatenreihe(anzRollen - 1)
        ReDim Xdatenreihe(anzRollen - 1)



        For r = 1 To anzRollen
            roleName = RoleDefinitions.getRoledef(r).name
            tdatenreihe(r - 1) = ShowProjekte.getAuslastungsValues(roleName, auswahl).Sum
            Xdatenreihe(r - 1) = roleName
        Next r


        If auswahl = 1 Then
            titelTeile(0) = portfolioDiagrammtitel(PTpfdk.UeberAuslastung) & " (" & awinSettings.kapaEinheit & ")"
        Else
            titelTeile(0) = portfolioDiagrammtitel(PTpfdk.Unterauslastung) & " (" & awinSettings.kapaEinheit & ")"
        End If


        titelTeilLaengen(0) = titelTeile(0).Length + 1
        titelTeile(1) = textZeitraum(showRangeLeft, showRangeRight)
        titelTeilLaengen(1) = titelTeile(1).Length
        diagramTitle = titelTeile(0) & vbLf & titelTeile(1)



        With appInstance.Worksheets(arrWsNames(3))


            Dim tmpValues(1) As Double



            With chtobj.Chart
                '' remove extra series
                'Do Until .SeriesCollection.Count = 0
                '    .SeriesCollection(1).Delete()
                'Loop


                ' -----------------------
                ' Schreibe Über- bzw Unterauslastung 

                With .SeriesCollection(1)
                    .name = "Details"

                    .Values = tdatenreihe
                    .XValues = Xdatenreihe

                    .ChartType = Excel.XlChartType.xlPie
                    .HasDataLabels = True

                    With .Datalabels
                        .Position = Excel.XlDataLabelPosition.xlLabelPositionOutsideEnd
                        ' ur: 17.7.2014 fontsize kommt vom existierenden chart
                        '.Font.Size = awinSettings.fontsizeItems + 2
                    End With

                End With


                For r = 1 To anzRollen

                    roleName = RoleDefinitions.getRoledef(r).name
                    With .SeriesCollection(1).Points(r)
                        .Interior.color = RoleDefinitions.getRoledef(roleName).farbe
                        ' ur: 17.7.2014 fontsize kommt vom existierenden chart
                        '.DataLabel.Font.Size = awinSettings.fontsizeItems
                    End With

                Next r

                .HasTitle = True
                .ChartTitle.Text = diagramTitle
                ' ur: 17.7.2014 fontsize kommt vom existierenden chart
                ' .ChartTitle.Font.Size = awinSettings.fontsizeTitle
                '.ChartTitle.Format.TextFrame2.TextRange.Characters(titelTeilLaengen(0) + 1, _
                '    titelTeilLaengen(1)).Font.Size = awinSettings.fontsizeLegend

            End With

        End With

    End Sub


    ''' <summary>
    ''' Prozedur aktualisiert das Kosten Struktur Chart des Projektes  (Balken-Diagramm)
    ''' </summary>
    ''' <param name="hproj">Verweis auf Projekt</param>
    ''' <param name="chtobj">Verweis auf ChartObject</param>
    ''' <param name="auswahl">
    ''' Auswahl = 1 : Diagramm zeigt nur sonstige Kosten 
    ''' Auswahl = 2 : Diagramm zeigt alle Kosten, inkl Personalkosten 
    ''' </param>
    ''' <remarks>
    ''' kennziffer fest auf 3 gesetzt 
    ''' kennziffer = 0 : Phasen Diagramm
    '''            = 1 : Personal-Bedarfe (Balken)
    '''            = 2 : Personal-Bedarfe (PIE)
    '''            = 3 : Kosten (Balken)
    '''            = 4 : Kosten (Pie)
    '''            = 5 : Strategie / Risiko 
    '''            = 6 : Ergebnis
    ''' </remarks>
    Public Sub updateCostBalkenOfProject(ByRef hproj As clsProjekt, ByRef chtobj As Excel.ChartObject, _
                                         ByVal auswahl As Integer, ByVal changeScale As Boolean)


        Dim kennziffer As Integer = 3
        Dim plen As Integer
        Dim i As Integer
        Dim Xdatenreihe() As String
        Dim tdatenreihe() As Double
        Dim sumdatenreihe() As Double
        Dim anzKostenarten As Integer
        Dim costname As String
        Dim pkIndex As Integer = CostDefinitions.Count
        Dim pstart As Integer
        Dim diagramTitle As String
        Dim titelTeile(1) As String
        Dim titelTeilLaengen(1) As Integer
        Dim tmpcollection As New Collection
        Dim kennung As String = " "


        Dim pname As String = hproj.name

        tmpcollection.Add(hproj.getShapeText & "#" & auswahl.ToString)
        kennung = calcChartKennung("pr", PTprdk.KostenBalken, tmpcollection)

        If auswahl = 1 Then
            titelTeile(0) = "Sonstige Kosten T€" & vbLf & hproj.getShapeText & vbLf
            titelTeilLaengen(0) = titelTeile(0).Length
            titelTeile(1) = " (" & hproj.timeStamp.ToString & ") "
            titelTeilLaengen(1) = titelTeile(1).Length
            diagramTitle = titelTeile(0) & titelTeile(1)

        Else
            titelTeile(0) = "Gesamtkosten T€" & vbLf & hproj.getShapeText & vbLf
            titelTeilLaengen(0) = titelTeile(0).Length
            titelTeile(1) = " (" & hproj.timeStamp.ToString & ") "
            titelTeilLaengen(1) = titelTeile(1).Length
            diagramTitle = titelTeile(0) & titelTeile(1)

        End If


        Dim ErgebnisListeK As Collection

        Dim formerEE As Boolean = appInstance.EnableEvents
        appInstance.EnableEvents = False

        '
        ' hole die Projektdauer
        '
        With hproj
            plen = .anzahlRasterElemente
            pstart = .Start
        End With
        '
        ' hole die Anzahl Kostenarten, die in diesem Projekt vorkommen
        '
        ErgebnisListeK = hproj.getUsedKosten
        anzKostenarten = ErgebnisListeK.Count

        'If anzKostenarten = 0 Then
        '    MsgBox("keine Kosten-Bedarfe definiert")
        '    Exit Sub
        'End If


        ReDim Xdatenreihe(plen - 1)
        ReDim tdatenreihe(plen - 1)
        ReDim sumdatenreihe(plen - 1)



        For i = 1 To plen
            Xdatenreihe(i - 1) = StartofCalendar.AddMonths(pstart + i - 2).ToString("MMM yy")
        Next i


        Dim ik As Integer = 1 ' wird für die Unterscheidung benötigt, ob mit Personal-Kosten oder ohne 


        With CType(chtobj.Chart, Excel.Chart)

            '' remove extra series
            Do Until .SeriesCollection.Count = 0
                .SeriesCollection(1).Delete()
            Loop

            If auswahl = 2 Then
                ik = 0
                costname = "Personalkosten"
                tdatenreihe = hproj.getAllPersonalKosten
                For i = 0 To plen - 1
                    sumdatenreihe(i) = sumdatenreihe(i) + tdatenreihe(i)
                Next

                With .SeriesCollection.NewSeries
                    .Name = costname
                    .Interior.color = CostDefinitions.getCostdef(pkIndex).farbe
                    .Values = tdatenreihe
                    .XValues = Xdatenreihe
                    '.ChartType = Excel.XlChartType.xlColumnStacked
                End With
            End If

            For k = 1 To anzKostenarten
                costname = CStr(ErgebnisListeK.Item(k))
                tdatenreihe = hproj.getKostenBedarf(costname)

                For i = 0 To plen - 1
                    sumdatenreihe(i) = sumdatenreihe(i) + tdatenreihe(i)
                Next
                Dim iSerColl As Integer = CType(.SeriesCollection, Excel.SeriesCollection).Count

                With .SeriesCollection.NewSeries
                    .name = costname
                    .Interior.color = CostDefinitions.getCostdef(costname).farbe
                    .Values = tdatenreihe
                    .XValues = Xdatenreihe
                    '.ChartType = Excel.XlChartType.xlColumnStacked
                End With

            Next k

            If CBool(.HasAxis(Excel.XlAxisType.xlValue)) Then

                With CType(.Axes(Excel.XlAxisType.xlValue), Excel.Axis)
                    ' das ist dann relevant, wenn ein anderes Projekt selektiert wird, das über die aktuelle Skalierung 
                    ' hinausgehende Werte hat 

                    If changeScale Then
                        .MinimumScale = 0
                        If Not (.MaximumScaleIsAuto) Then

                            If sumdatenreihe.Max > .MaximumScale - 3 Then
                                .MaximumScale = sumdatenreihe.Max + 3
                            End If
                            .MaximumScaleIsAuto = True
                        End If
                    End If


                End With

            End If

            If .HasTitle Then
                .ChartTitle.Text = diagramTitle
                ' ur: 21.07.2014 für Chart-Cockpit auskommentiert
                '.ChartTitle.Font.Size = awinSettings.fontsizeTitle
                '.ChartTitle.Format.TextFrame2.TextRange.Characters(titelTeilLaengen(0) + 1, _
                '        titelTeilLaengen(1)).Font.Size = awinSettings.fontsizeLegend
            End If

        End With

        chtobj.Name = kennung

        appInstance.EnableEvents = formerEE
        'appInstance.ScreenUpdating = formerSU


    End Sub


    ''' <summary>
    ''' zeigt die Zusammensetzung der Überauslastung bzw Unterauslastung an 
    ''' 
    ''' </summary>
    ''' <param name="repObj">Verweis auf das erzeugte Chart</param>
    ''' <param name="auswahl">
    ''' 1 = Überauslastung
    ''' 2 = Unterauslastung
    ''' </param>
    ''' <param name="top">Diagramm Koordinate oben</param>
    ''' <param name="left">Diagramm Koordinate links</param>
    ''' <param name="height">Diagramm-Höhe</param>
    ''' <param name="width">Diagramm-Breite</param>
    ''' <remarks></remarks>
    Public Sub createAuslastungsDetailPie(ByRef repObj As Excel.ChartObject, ByVal auswahl As Integer, _
                                                ByVal top As Double, left As Double, height As Double, width As Double, _
                                                ByVal calledfromReporting As Boolean)


        Dim diagramTitle As String
        Dim anzDiagrams As Integer
        Dim chtobjname As String



        Dim Xdatenreihe() As String
        Dim tdatenreihe() As Double

        Dim anzRollen As Integer
        Dim roleName As String


        Dim kennung As String
        Dim zE As String = awinSettings.kapaEinheit
        Dim titelTeile(1) As String
        Dim titelTeilLaengen(1) As Integer

        Dim myCollection As New Collection
        myCollection.Add("Auslastungs-Details")

        If auswahl = 1 Then
            chtobjname = calcChartKennung("pf", PTpfdk.UeberAuslastung, myCollection)
        Else
            chtobjname = calcChartKennung("pf", PTpfdk.Unterauslastung, myCollection)
        End If
        myCollection.Clear()

        If Not calledfromReporting Then

            Dim foundDiagramm As clsDiagramm

            ' wenn die Werte für dieses Diagramm bereits einmal gespeichert wurden ... -> übernehmen 
            Try
                foundDiagramm = DiagramList.getDiagramm(chtobjname)
                With foundDiagramm
                    top = .top
                    left = .left
                    width = .width
                    height = .height
                End With
            Catch ex As Exception


            End Try
        End If



        Dim formerEE As Boolean = appInstance.EnableEvents
        appInstance.EnableEvents = False


        '
        ' hole die Anzahl Rollen
        '
        anzRollen = RoleDefinitions.Count

        If anzRollen = 0 Then
            MsgBox("keine Rollen-Bedarfe definiert")
            Exit Sub
        End If

        ReDim tdatenreihe(anzRollen - 1)

        ReDim Xdatenreihe(anzRollen - 1)



        For r = 1 To anzRollen
            roleName = RoleDefinitions.getRoledef(r).name
            tdatenreihe(r - 1) = ShowProjekte.getAuslastungsValues(roleName, auswahl).Sum
            Xdatenreihe(r - 1) = roleName
        Next r


        If auswahl = 1 Then
            titelTeile(0) = portfolioDiagrammtitel(PTpfdk.UeberAuslastung) & " (" & awinSettings.kapaEinheit & ")"
        Else
            titelTeile(0) = portfolioDiagrammtitel(PTpfdk.Unterauslastung) & " (" & awinSettings.kapaEinheit & ")"
        End If

        titelTeilLaengen(0) = titelTeile(0).Length + 1
        titelTeile(1) = textZeitraum(showRangeLeft, showRangeRight)
        titelTeilLaengen(1) = titelTeile(1).Length
        diagramTitle = titelTeile(0) & vbLf & titelTeile(1)
        kennung = titelTeile(0)



        With CType(appInstance.Worksheets(arrWsNames(3)), Excel.Worksheet)
            anzDiagrams = CType(.ChartObjects, Excel.ChartObjects).Count
            '
            ' um welches Diagramm handelt es sich ...
            '
            Dim i As Integer = 1
            Dim found As Boolean = False

            While i <= anzDiagrams And Not found

                If .ChartObjects(i).name = chtobjname Then
                    found = True
                Else
                    i = i + 1
                End If

            End While

            If found Then
                'Call MsgBox("Chart wird bereits angezeigt ...")
                appInstance.EnableEvents = formerEE
                repObj = CType(.ChartObjects(i), Excel.ChartObject)
                'appInstance.ScreenUpdating = formerSU
                Exit Sub
            Else
                Dim tmpValues(1) As Double



                With appInstance.Charts.Add
                    ' remove extra series
                    Do Until .SeriesCollection.Count = 0
                        .SeriesCollection(1).Delete()
                    Loop


                    ' -----------------------
                    ' Schreibe Über- bzw Unterauslastung 

                    With .SeriesCollection.NewSeries
                        .name = "Details"

                        .Values = tdatenreihe
                        .XValues = Xdatenreihe

                        .ChartType = Excel.XlChartType.xlPie
                        .HasDataLabels = True

                        With .Datalabels
                            .Position = Excel.XlDataLabelPosition.xlLabelPositionOutsideEnd
                            .Font.Size = awinSettings.fontsizeItems + 2
                        End With

                    End With


                    For r = 1 To anzRollen

                        roleName = RoleDefinitions.getRoledef(r).name
                        With .SeriesCollection(1).Points(r)
                            .Interior.color = RoleDefinitions.getRoledef(roleName).farbe
                            .DataLabel.Font.Size = awinSettings.fontsizeItems
                        End With

                    Next r


                    .HasLegend = True
                    With .Legend
                        .Position = Excel.Constants.xlRight
                        .Font.Size = awinSettings.fontsizeItems + 2
                    End With

                    .HasTitle = True
                    .ChartTitle.text = diagramTitle
                    .ChartTitle.Font.Size = awinSettings.fontsizeTitle
                    .ChartTitle.Format.TextFrame2.TextRange.Characters(titelTeilLaengen(0) + 1, _
                        titelTeilLaengen(1)).Font.Size = awinSettings.fontsizeLegend
                    .Location(Where:=XlChartLocation.xlLocationAsObject, Name:=appInstance.Worksheets(arrWsNames(3)).name)
                End With



                With .ChartObjects(anzDiagrams + 1)
                    .Name = chtobjname
                    .top = top
                    .left = left
                    .height = height
                    .width = width
                End With
            End If


            repObj = CType(.ChartObjects(anzDiagrams + 1), Excel.ChartObject)

            ' jetzt muss die letzte Position des Diagramms gespeichert werden , wenn es nicht aus der Reporting Engine 
            ' aufgerufen wurde
            If Not calledfromReporting Then

                Dim prcDiagram As New clsDiagramm

                ' Anfang Event Handling für Chart 
                Dim prcChart As New clsEventsPrcCharts
                prcChart.PrcChartEvents = CType(.ChartObjects(anzDiagrams + 1), Excel.ChartObject).Chart
                prcDiagram.setDiagramEvent = prcChart
                ' Ende Event Handling für Chart 


                With prcDiagram
                    .DiagrammTitel = diagramTitle
                    .diagrammTyp = DiagrammTypen(4)
                    .gsCollection = Nothing
                    .isCockpitChart = False
                    .top = top
                    .left = left
                    .width = width
                    .height = height
                    .kennung = chtobjname
                End With

                ' eintragen in die sortierte Liste mit .kennung als dem Schlüssel 
                ' wenn das Diagramm bereits existiert, muss es gelöscht werden, dann neu ergänzt ... 
                Try
                    DiagramList.Add(prcDiagram)
                Catch ex As Exception

                    Try
                        DiagramList.Remove(prcDiagram.kennung)
                        DiagramList.Add(prcDiagram)
                    Catch ex1 As Exception

                    End Try


                End Try

            End If





        End With




    End Sub




    '
    ' Prozedur zeigt die Kosten Struktur des Projektes an (Balken-Diagramm)
    '
    ' Auswahl = 1 : Diagramm zeigt nur sonstige Kosten 
    ' Auswahl = 2 : Diagramm zeigt alle Kosten, inkl Personalkosten 
    ' kennziffer = 0 : Phasen Diagramm
    '            = 1 : Personal-Bedarfe (Balken)
    '            = 2 : Personal-Bedarfe (PIE)
    '            = 3 : Kosten (Balken)
    '            = 4 : Kosten (Pie)
    '            = 5 : Strategie / Risiko 
    '            = 6 : Ergebnis

    Public Sub createRessPieOfProject(ByRef hproj As clsProjekt, ByRef repObj As Excel.ChartObject, ByVal auswahl As Integer, _
                                        ByVal top As Double, left As Double, height As Double, width As Double)

        'Dim kennziffer As Integer = 4
        Dim diagramTitle As String
        Dim anzDiagrams As Integer

        Dim plen As Integer

        Dim Xdatenreihe() As String
        Dim tdatenreihe() As Double

        Dim anzRollen As Integer
        Dim roleName As String

        'Dim prIndex As Integer = RoleDefinitions.Count
        Dim pstart As Integer
        Dim pname As String = hproj.name

        Dim kennung As String
        Dim zE As String = awinSettings.kapaEinheit
        Dim titelTeile(1) As String
        Dim titelTeilLaengen(1) As Integer
        Dim tmpcollection As New Collection


        Dim ErgebnisListeR As Collection
        Dim formerEE As Boolean = appInstance.EnableEvents
        appInstance.EnableEvents = False


        '
        ' hole die Projektdauer
        '
        With hproj
            plen = .anzahlRasterElemente
            pstart = .Start
        End With
        '
        ' hole die Anzahl Kostenarten, die in diesem Projekt vorkommen
        '
        ErgebnisListeR = hproj.getUsedRollen
        anzRollen = ErgebnisListeR.Count

        If anzRollen = 0 Then
            MsgBox("keine Rollen-Bedarfe definiert")
            Exit Sub
        End If

        ReDim tdatenreihe(anzRollen - 1)
        ReDim Xdatenreihe(anzRollen - 1)


        For r = 0 To anzRollen - 1
            roleName = CStr(ErgebnisListeR.Item(r + 1))
            Xdatenreihe(r) = roleName
            If auswahl = 1 Then
                tdatenreihe(r) = Math.Round(hproj.getRessourcenBedarf(roleName).Sum)
            Else
                tdatenreihe(r) = Math.Round(hproj.getPersonalKosten(roleName).Sum)
            End If

        Next r

        tmpcollection.Add(hproj.getShapeText & "#" & auswahl.ToString)
        kennung = calcChartKennung("pr", PTprdk.PersonalPie, tmpcollection)

        If auswahl = 1 Then
            titelTeile(0) = "Personalbedarf (" & tdatenreihe.Sum.ToString("#####.") & zE & ")" & vbLf & hproj.getShapeText & vbLf
            titelTeilLaengen(0) = titelTeile(0).Length
            titelTeile(1) = "(" & hproj.timeStamp.ToString & ") "
            titelTeilLaengen(1) = titelTeile(1).Length
            diagramTitle = titelTeile(0) & titelTeile(1)
            'kennung = "Personalbedarf"
        Else
            titelTeile(0) = "Personalkosten (" & tdatenreihe.Sum.ToString("#####.") & " T€)" & vbLf & hproj.getShapeText & vbLf
            titelTeilLaengen(0) = titelTeile(0).Length
            titelTeile(1) = "(" & hproj.timeStamp.ToString & ") "
            titelTeilLaengen(1) = titelTeile(1).Length
            diagramTitle = titelTeile(0) & titelTeile(1)

            'kennung = "Personalkosten"
        End If


        With CType(appInstance.Worksheets(arrWsNames(3)), Excel.Worksheet)
            anzDiagrams = CType(.ChartObjects, Excel.ChartObjects).Count
            '
            ' um welches Diagramm handelt es sich ...
            '
            Dim i As Integer = 1
            Dim found As Boolean = False
            Dim chtTitle As String
            While i <= anzDiagrams And Not found
                Try
                    chtTitle = CType(.ChartObjects(i), Excel.ChartObject).Chart.ChartTitle.Text
                Catch ex As Exception
                    chtTitle = " "
                End Try

                If chtTitle = diagramTitle Then
                    found = True

                Else
                    i = i + 1
                End If

            End While

            If found Then
                'Call MsgBox("Chart wird bereits angezeigt ...")
                appInstance.EnableEvents = formerEE
                repObj = CType(.ChartObjects(i), Excel.ChartObject)
                'appInstance.ScreenUpdating = formerSU
                Exit Sub
            Else
                With appInstance.Charts.Add
                    ' remove extra series
                    Do Until .SeriesCollection.Count = 0
                        .SeriesCollection(1).Delete()
                    Loop

                    With .SeriesCollection.NewSeries
                        .name = pname
                        .Values = tdatenreihe
                        .XValues = Xdatenreihe
                        .ChartType = Excel.XlChartType.xlPie
                        .HasDataLabels = True
                        .DataLabels.Position = Excel.XlDataLabelPosition.xlLabelPositionOutsideEnd
                    End With

                    For r = 1 To anzRollen
                        roleName = CStr(ErgebnisListeR.Item(r))
                        With .SeriesCollection(1).Points(r)
                            .Interior.color = RoleDefinitions.getRoledef(roleName).farbe
                            .DataLabel.Font.Size = awinSettings.fontsizeItems
                        End With
                    Next r

                    .HasLegend = True
                    With .Legend
                        .Position = Excel.Constants.xlRight
                        .Font.Size = awinSettings.fontsizeItems
                    End With
                    .HasTitle = True
                    .ChartTitle.text = diagramTitle
                    .ChartTitle.Font.Size = awinSettings.fontsizeTitle
                    .ChartTitle.Format.TextFrame2.TextRange.Characters(titelTeilLaengen(0) + 1, _
                        titelTeilLaengen(1)).Font.Size = awinSettings.fontsizeLegend
                    .Location(Where:=XlChartLocation.xlLocationAsObject, Name:=appInstance.Worksheets(arrWsNames(3)).name)
                End With

                With .ChartObjects(anzDiagrams + 1)
                    '.Name = pname & "#" & kennung & "#" & "2"
                    .Name = kennung
                    .top = top
                    .left = left
                    .height = height
                    .width = width
                End With
            End If


            repObj = CType(.ChartObjects(anzDiagrams + 1), Excel.ChartObject)


        End With




    End Sub

    '
    ' Prozedur zeigt die Kosten Struktur des Projektes an (Balken-Diagramm)
    '
    ' Auswahl = 1 : Diagramm zeigt nur sonstige Kosten 
    ' Auswahl = 2 : Diagramm zeigt alle Kosten, inkl Personalkosten 
    ' kennziffer = 0 : Phasen Diagramm
    '            = 1 : Personal-Bedarfe (Balken)
    '            = 2 : Personal-Bedarfe (PIE)
    '            = 3 : Kosten (Balken)
    '            = 4 : Kosten (Pie)
    '            = 5 : Strategie / Risiko 
    '            = 6 : Ergebnis

    Public Sub updateRessPieOfProject(ByRef hproj As clsProjekt, ByRef chtobj As Excel.ChartObject, ByVal auswahl As Integer)

        'Dim kennziffer As Integer = 4
        Dim diagramTitle As String
        'Dim anzDiagrams As Integer

        Dim plen As Integer

        Dim Xdatenreihe() As String
        Dim tdatenreihe() As Double

        Dim anzRollen As Integer
        Dim roleName As String


        Dim pstart As Integer
        Dim pname As String = hproj.name


        Dim kennung As String = " "
        Dim zE As String = awinSettings.kapaEinheit & " "
        Dim titelTeile(1) As String
        Dim titelTeilLaengen(1) As Integer
        Dim tmpCollection As New Collection


        Dim ErgebnisListeR As Collection


        Dim formerEE As Boolean = appInstance.EnableEvents
        appInstance.EnableEvents = False


        '
        ' hole die Projektdauer
        '
        With hproj
            plen = .anzahlRasterElemente
            pstart = .Start
        End With
        '
        ' hole die Anzahl Kostenarten, die in diesem Projekt vorkommen
        '
        ErgebnisListeR = hproj.getUsedRollen
        anzRollen = ErgebnisListeR.Count

        ' sonst kommt der in eine Endlos Schleife, wenn keine Rollen definiert sind 
        If anzRollen > 0 Then
            ReDim tdatenreihe(anzRollen - 1)
            ReDim Xdatenreihe(anzRollen - 1)
        Else
            ReDim tdatenreihe(0)
            ReDim Xdatenreihe(0)
        End If

        


        For r = 0 To anzRollen - 1
            roleName = CStr(ErgebnisListeR.Item(r + 1))
            Xdatenreihe(r) = roleName

            If auswahl = 1 Then
                tdatenreihe(r) = Math.Round(hproj.getRessourcenBedarf(roleName).Sum)
            Else
                tdatenreihe(r) = Math.Round(hproj.getPersonalKosten(roleName).Sum / 10) * 10
            End If

        Next r

        tmpCollection.Add(hproj.getShapeText & "#" & auswahl.ToString)
        kennung = calcChartKennung("pr", PTprdk.PersonalPie, tmpCollection)

        If auswahl = 1 Then
            titelTeile(0) = "Personalbedarf (" & tdatenreihe.Sum.ToString("####.#") & zE & ")" & vbLf & hproj.getShapeText & vbLf
            titelTeilLaengen(0) = titelTeile(0).Length
            titelTeile(1) = "(" & hproj.timeStamp.ToString & ") "
            titelTeilLaengen(1) = titelTeile(1).Length
            diagramTitle = titelTeile(0) & titelTeile(1)

            'kennung = "Personalbedarf"
        Else
            titelTeile(0) = "Personalkosten (" & tdatenreihe.Sum.ToString("####.#") & " T€)" & vbLf & hproj.getShapeText & vbLf
            titelTeilLaengen(0) = titelTeile(0).Length
            titelTeile(1) = "(" & hproj.timeStamp.ToString & ") "
            titelTeilLaengen(1) = titelTeile(1).Length
            diagramTitle = titelTeile(0) & titelTeile(1)

            'kennung = "Personalkosten"
        End If



        With chtobj.Chart
            'ur:22.07.2014 wegen Chart-Cockpit
            '' remove extra series
            'Do Until .SeriesCollection.Count = 0
            '    .SeriesCollection(1).Delete()
            'Loop

            'With .SeriesCollection.NewSeries
            With .SeriesCollection(1)
                .Name = pname
                .Values = tdatenreihe
                .XValues = Xdatenreihe
                .ChartType = Excel.XlChartType.xlPie
                .HasDataLabels = True
                .DataLabels.Position = Excel.XlDataLabelPosition.xlLabelPositionOutsideEnd
            End With

            For r = 1 To anzRollen
                roleName = CStr(ErgebnisListeR.Item(r))
                With .SeriesCollection(1).Points(r)
                    .Interior.color = RoleDefinitions.getRoledef(roleName).farbe
                    ' ur: 21.07.2014 für Chart-Cockpit auskommentiert
                    '.DataLabel.Font.Size = awinSettings.fontsizeItems
                End With
            Next r

            ' Änderung: evtl wurde ja der Titel gelöscht 
            If .HasTitle Then
                .ChartTitle.Text = diagramTitle
                ' ur: 21.07.2014 für Chart-Cockpit auskommentiert
                '.ChartTitle.Font.Size = awinSettings.fontsizeTitle
                '.ChartTitle.Format.TextFrame2.TextRange.Characters(titelTeilLaengen(0) + 1, _
                '        titelTeilLaengen(1)).Font.Size = awinSettings.fontsizeLegend
            End If


        End With

        chtobj.Name = kennung

        appInstance.EnableEvents = formerEE



    End Sub
    '
    ' Prozedur zeigt die Kosten Struktur des Projektes an (Balken-Diagramm)
    '
    ' Auswahl = 1 : Diagramm zeigt nur sonstige Kosten 
    ' Auswahl = 2 : Diagramm zeigt alle Kosten, inkl Personalkosten 
    ' kennziffer = 0 : Phasen Diagramm
    '            = 1 : Personal-Bedarfe (Balken)
    '            = 2 : Personal-Bedarfe (PIE)
    '            = 3 : Kosten (Balken)
    '            = 4 : Kosten (Pie)
    '            = 5 : Strategie / Risiko 
    '            = 6 : Ergebnis

    Public Sub createCostPieOfProject(ByRef hproj As clsProjekt, ByRef repObj As Excel.ChartObject, ByVal auswahl As Integer, _
                                        ByVal top As Double, left As Double, height As Double, width As Double)

        Dim kennziffer As Integer = 4
        Dim diagramTitle As String
        Dim anzDiagrams As Integer

        Dim plen As Integer

        Dim Xdatenreihe() As String
        Dim tdatenreihe() As Double

        Dim anzKostenarten As Integer
        Dim costname As String

        Dim pkIndex As Integer = CostDefinitions.Count
        Dim pstart As Integer
        Dim pname As String = hproj.name

        Dim kennung As String
        Dim titelTeile(1) As String
        Dim titelTeilLaengen(1) As Integer
        Dim tmpcollection As New Collection


        Dim ErgebnisListeK As Collection

        Dim formerEE As Boolean = appInstance.EnableEvents
        appInstance.EnableEvents = False
        'appInstance.ScreenUpdating = False


        '
        ' hole die Projektdauer
        '
        With hproj
            plen = .anzahlRasterElemente
            pstart = .Start
        End With
        '
        ' hole die Anzahl Kostenarten, die in diesem Projekt vorkommen
        '
        ErgebnisListeK = hproj.getUsedKosten
        anzKostenarten = ErgebnisListeK.Count

        ' Änderung: es wird die Null gezeigt
        'If anzKostenarten = 0 And auswahl = 1 Then
        '    appInstance.EnableEvents = formerEE
        '    Throw New Exception("keine Kosten-Bedarfe definiert")
        'End If

        tmpcollection.Add(hproj.getShapeText & "#" & auswahl.ToString)
        kennung = calcChartKennung("pr", PTprdk.KostenPie, tmpcollection)

        If auswahl = 1 Then
            ' Alle Sonstigen Kostenarten 
            If anzKostenarten = 0 Then
                ReDim tdatenreihe(0)
                ReDim Xdatenreihe(0)
            Else
                ReDim tdatenreihe(anzKostenarten - 1)
                ReDim Xdatenreihe(anzKostenarten - 1)
            End If

        Else
            ' alle Kostenarten - inkl Personalkosten 
            ReDim tdatenreihe(anzKostenarten)
            ReDim Xdatenreihe(anzKostenarten)
            'Xdatenreihe(0) = "Personal-Kosten"
            'tdatenreihe(0) = hsum(0)
        End If


        For k = 0 To anzKostenarten - 1
            costname = CStr(ErgebnisListeK.Item(k + 1))
            Xdatenreihe(k) = costname
            tdatenreihe(k) = Math.Round(hproj.getKostenBedarf(costname).Sum)
        Next k

        If auswahl = 2 Then
            Xdatenreihe(anzKostenarten) = "Personal-Kosten"
            tdatenreihe(anzKostenarten) = Math.Round(hproj.getAllPersonalKosten.Sum)
        End If

        If auswahl = 1 Then
            titelTeile(0) = "Sonstige Kosten (" & tdatenreihe.Sum.ToString("#####.") & " T€)" & vbLf & hproj.getShapeText & vbLf
            titelTeilLaengen(0) = titelTeile(0).Length
            titelTeile(1) = " (" & hproj.timeStamp.ToString & ") "
            titelTeilLaengen(1) = titelTeile(1).Length
            diagramTitle = titelTeile(0) & titelTeile(1)

            'kennung = "Sonstige Kosten"
        Else
            titelTeile(0) = "Gesamtkosten (" & tdatenreihe.Sum.ToString("#####.") & " T€)" & vbLf & hproj.getShapeText & vbLf
            titelTeilLaengen(0) = titelTeile(0).Length
            titelTeile(1) = " (" & hproj.timeStamp.ToString & ") "
            titelTeilLaengen(1) = titelTeile(1).Length
            diagramTitle = titelTeile(0) & titelTeile(1)

            'kennung = "Gesamtkosten"
        End If

        If tdatenreihe.Sum = 0.0 Then
            appInstance.EnableEvents = formerEE
            Throw New Exception("Summe sonstige Kosten ist Null")
        Else
            With CType(appInstance.Worksheets(arrWsNames(3)), Excel.Worksheet)
                anzDiagrams = CType(.ChartObjects, Excel.ChartObjects).Count

                '
                ' um welches Diagramm handelt es sich ...
                '
                Dim i As Integer = 1
                Dim found As Boolean = False
                Dim chtTitle As String
                While i <= anzDiagrams And Not found
                    Try
                        chtTitle = CType(.ChartObjects(i), Excel.ChartObject).Chart.ChartTitle.Text
                    Catch ex As Exception
                        chtTitle = " "
                    End Try

                    If chtTitle = diagramTitle Then
                        found = True

                    Else
                        i = i + 1
                    End If

                End While

                If found Then
                    'Call MsgBox("Chart wird bereits angezeigt ...")
                    repObj = CType(.ChartObjects(i), Excel.ChartObject)
                    'appInstance.ScreenUpdating = formerSU
                Else
                    With appInstance.Charts.Add
                        ' remove extra series
                        Do Until .SeriesCollection.Count = 0
                            .SeriesCollection(1).Delete()
                        Loop

                        With .SeriesCollection.NewSeries
                            .name = pname
                            .Values = tdatenreihe
                            .XValues = Xdatenreihe
                            .ChartType = Excel.XlChartType.xlPie
                            .HasDataLabels = True
                            .DataLabels.Position = Excel.XlDataLabelPosition.xlLabelPositionOutsideEnd
                        End With

                        For k = 0 To anzKostenarten - 1 + auswahl - 1
                            If k = anzKostenarten Then
                                costname = "Personal-Kosten"
                                With .SeriesCollection(1).Points(k + 1)
                                    .Interior.color = CostDefinitions.getCostdef(pkIndex).farbe
                                    .DataLabel.Font.Size = 10

                                End With
                            Else
                                costname = CStr(ErgebnisListeK.Item(k + 1))
                                With .SeriesCollection(1).Points(k + 1)
                                    .Interior.color = CostDefinitions.getCostdef(costname).farbe
                                    .DataLabel.Font.Size = 10

                                End With
                            End If

                        Next k

                        .HasLegend = True
                        With .Legend
                            .Position = Excel.Constants.xlRight
                            .Font.Size = awinSettings.fontsizeItems
                        End With
                        .HasTitle = True
                        .ChartTitle.text = diagramTitle
                        .ChartTitle.Font.Size = awinSettings.fontsizeTitle
                        .ChartTitle.Format.TextFrame2.TextRange.Characters(titelTeilLaengen(0) + 1, _
                                titelTeilLaengen(1)).Font.Size = awinSettings.fontsizeLegend
                        .Location(Where:=XlChartLocation.xlLocationAsObject, Name:=appInstance.Worksheets(arrWsNames(3)).name)
                    End With

                    With .ChartObjects(anzDiagrams + 1)
                        '.Name = pname & "#" & kennung & "#" & "2"
                        .Name = kennung
                        .top = top
                        .left = left
                        .height = height
                        .width = width
                    End With

                    repObj = CType(.ChartObjects(anzDiagrams + 1), Excel.ChartObject)
                End If



            End With

        End If


        appInstance.EnableEvents = formerEE
        'appInstance.ScreenUpdating = True


    End Sub

   
    ''' <summary>
    ''' aktualisiert im exist. Chart die Kosten Struktur des Projektes (Pie-Chart)
    ''' Kennziffer fest = 4 (Kosten Pie)
    ''' </summary>
    ''' <param name="hproj">Verweis auf das Projekt</param>
    ''' <param name="chtobj">Verweis auf das Chart-Objekt</param>
    ''' <param name="auswahl">
    ''' Auswahl = 1 : Diagramm zeigt nur sonstige Kosten 
    ''' Auswahl = 2 : Diagramm zeigt alle Kosten, inkl Personalkosten 
    ''' </param>
    ''' <remarks>
    ''' kennziffer = 0 : Phasen Diagramm
    '''            = 1 : Personal-Bedarfe (Balken)
    '''            = 2 : Personal-Bedarfe (PIE)
    '''            = 3 : Kosten (Balken)
    '''            = 4 : Kosten (Pie)
    '''            = 5 : Strategie / Risiko 
    '''            = 6 : Ergebnis
    ''' </remarks>
    Public Sub updateCostPieOfProject(ByRef hproj As clsProjekt, ByRef chtobj As Excel.ChartObject, ByVal auswahl As Integer)

        Dim kennziffer As Integer = 4
        Dim diagramTitle As String

        Dim plen As Integer

        Dim Xdatenreihe() As String
        Dim tdatenreihe() As Double

        Dim anzKostenarten As Integer
        Dim costname As String

        Dim pkIndex As Integer = CostDefinitions.Count
        Dim pstart As Integer
        Dim pname As String = hproj.name

        Dim kennung As String = " "
        Dim titelTeile(1) As String
        Dim titelTeilLaengen(1) As Integer
        Dim tmpCollection As New Collection



        Dim ErgebnisListeK As Collection


        Dim formerEE As Boolean = appInstance.EnableEvents
        appInstance.EnableEvents = False
        'appInstance.ScreenUpdating = False




        '
        ' hole die Projektdauer
        '
        With hproj
            plen = .anzahlRasterElemente
            pstart = .Start
        End With
        '
        ' hole die Anzahl Kostenarten, die in diesem Projekt vorkommen
        '
        ErgebnisListeK = hproj.getUsedKosten
        anzKostenarten = ErgebnisListeK.Count


        tmpCollection.Add(hproj.getShapeText & "#" & auswahl.ToString)
        kennung = calcChartKennung("pr", PTprdk.KostenPie, tmpCollection)

        If auswahl = 1 Then
            ' Alle Sonstigen Kostenarten 
            If anzKostenarten > 0 Then
                ReDim tdatenreihe(anzKostenarten - 1)
                ReDim Xdatenreihe(anzKostenarten - 1)
            Else
                ReDim tdatenreihe(0)
                ReDim Xdatenreihe(0)
            End If
           
        Else
            ' alle Kostenarten - inkl Personalkosten 
            ReDim tdatenreihe(anzKostenarten)
            ReDim Xdatenreihe(anzKostenarten)
            'Xdatenreihe(0) = "Personal-Kosten"
            'tdatenreihe(0) = hsum(0)
        End If


        For k = 0 To anzKostenarten - 1
            costname = CStr(ErgebnisListeK.Item(k + 1))
            Xdatenreihe(k) = costname
            tdatenreihe(k) = hproj.getKostenBedarf(costname).Sum
        Next k

        If auswahl = 2 Then
            Xdatenreihe(anzKostenarten) = "Personal-Kosten"
            tdatenreihe(anzKostenarten) = hproj.getAllPersonalKosten.Sum
        End If

        If auswahl = 1 Then
            titelTeile(0) = "Sonstige Kosten (" & tdatenreihe.Sum.ToString("####.#") & " T€)" & vbLf & hproj.getShapeText & vbLf
            titelTeilLaengen(0) = titelTeile(0).Length
            titelTeile(1) = " (" & hproj.timeStamp.ToString & ") "
            titelTeilLaengen(1) = titelTeile(1).Length
            diagramTitle = titelTeile(0) & titelTeile(1)

        Else
            titelTeile(0) = "Gesamtkosten (" & tdatenreihe.Sum.ToString("####.#") & " T€)" & vbLf & hproj.getShapeText & vbLf
            titelTeilLaengen(0) = titelTeile(0).Length
            titelTeile(1) = " (" & hproj.timeStamp.ToString & ") "
            titelTeilLaengen(1) = titelTeile(1).Length
            diagramTitle = titelTeile(0) & titelTeile(1)


        End If


        With chtobj.Chart
            ' ur: 22.07.2014: ist bereits im Cockpit-Chart enthalten
            '' remove extra series
            'Do Until .SeriesCollection.Count = 0
            '    .SeriesCollection(1).Delete()
            'Loop


            'With .SeriesCollection.NewSeries
            With .SeriesCollection(1)
                .Name = pname
                .Values = tdatenreihe
                .XValues = Xdatenreihe
                .ChartType = Excel.XlChartType.xlPie
                .HasDataLabels = True
                .DataLabels.Position = Excel.XlDataLabelPosition.xlLabelPositionOutsideEnd
            End With

            For k = 0 To anzKostenarten - 1 + auswahl - 1
                If k = anzKostenarten Then
                    costname = "Personal-Kosten"
                    With .SeriesCollection(1).Points(k + 1)
                        .Interior.color = CostDefinitions.getCostdef(pkIndex).farbe
                        ' ur: 21.07.2014 für Chart-Cockpit auskommentiert
                        '.DataLabel.Font.Size = 10

                    End With
                Else
                    costname = CStr(ErgebnisListeK.Item(k + 1))
                    With .SeriesCollection(1).Points(k + 1)
                        .Interior.color = CostDefinitions.getCostdef(costname).farbe
                        ' ur: 21.07.2014 für Chart-Cockpit auskommentiert
                        '.DataLabel.Font.Size = 10

                    End With
                End If

            Next k

            If .HasTitle Then
                .ChartTitle.Text = diagramTitle
                ' ur: 21.07.2014 für Chart-Cockpit auskommentiert
                '.ChartTitle.Font.Size = awinSettings.fontsizeTitle
                '.ChartTitle.Format.TextFrame2.TextRange.Characters(titelTeilLaengen(0) + 1, _
                '        titelTeilLaengen(1)).Font.Size = awinSettings.fontsizeLegend
            End If


        End With

        chtobj.Name = kennung

        appInstance.EnableEvents = formerEE



    End Sub

   
    ''' <summary>
    ''' zeigt die Entwicklung von Werten wie Personal-Kosten, Sonstige Kosten, Budget und Ergebnis an über die gesamte Projekt-Historie an 
    ''' Voraussetzung: die Projekt-Historie ist bestimmt 
    ''' </summary>
    ''' <param name="repObj">Verweis auf das erzeugte Chart</param>
    ''' <param name="top">Diagramm-Koordinate oben</param>
    ''' <param name="left">Diagramm-Koordinate links</param>
    ''' <param name="height">Diagramm-Höhe</param>
    ''' <param name="width">Diagramm-Breite</param>
    ''' <remarks></remarks>
    Public Sub createTrendKPI(ByRef repObj As Excel.ChartObject, ByVal top As Double, left As Double, height As Double, width As Double)

        Dim diagramTitle As String = " "
        Dim anzDiagrams As Integer
        Dim found As Boolean

        Dim i As Integer
        Dim Xdatenreihe() As String


        Dim chtTitle As String
        Dim pkIndex As Integer = CostDefinitions.Count

        Dim chtobj As Excel.ChartObject
        Dim ErgebnisListeR As New Collection

        Dim nrOFSnapshots As Integer = projekthistorie.Count

        If nrOFSnapshots = 0 Then
            Call MsgBox("es gibt keine Historie ...")
        End If

        Dim erloes() As Double
        Dim personalKosten() As Double
        Dim sonstKosten() As Double
        Dim risikoKosten() As Double
        Dim estimProfit() As Double



        Dim formerEE As Boolean = appInstance.EnableEvents
        'Dim formerSU As Boolean = appInstance.ScreenUpdating
        appInstance.EnableEvents = False
        'appInstance.ScreenUpdating = False

        ReDim erloes(nrOFSnapshots)
        ReDim personalKosten(nrOFSnapshots)
        ReDim sonstKosten(nrOFSnapshots)
        ReDim risikoKosten(nrOFSnapshots)
        ReDim estimProfit(nrOFSnapshots)
        ReDim Xdatenreihe(nrOFSnapshots)


        Dim pname As String = projekthistorie.Last.name

        diagramTitle = "Planungs-Historie Kennzahlen " & vbLf & projekthistorie.Last.getShapeText


        ' jetzt werden die einzelnen Werte aufgefüllt
        Dim ix As Integer = 0
        For Each kvp As KeyValuePair(Of Date, clsProjekt) In projekthistorie.liste

            With kvp.Value
                erloes(ix) = .Erloes
                personalKosten(ix) = .getAllPersonalKosten.Sum
                sonstKosten(ix) = .getGesamtAndereKosten.Sum
                risikoKosten(ix) = .risikoKostenfaktor * (personalKosten(ix) + sonstKosten(ix))
                estimProfit(ix) = erloes(ix) - (personalKosten(ix) + sonstKosten(ix) + risikoKosten(ix))
                Xdatenreihe(ix) = .timeStamp.ToString("d")
            End With

            ix = ix + 1

        Next

        Dim hproj As clsProjekt = ShowProjekte.getProject(pname)
        With hproj
            erloes(nrOFSnapshots) = .Erloes
            personalKosten(nrOFSnapshots) = .getAllPersonalKosten.Sum
            sonstKosten(nrOFSnapshots) = .getGesamtAndereKosten.Sum
            risikoKosten(nrOFSnapshots) = .risikoKostenfaktor * (personalKosten(ix) + sonstKosten(ix))
            estimProfit(nrOFSnapshots) = erloes(ix) - (personalKosten(ix) + sonstKosten(ix) + risikoKosten(ix))
            Xdatenreihe(nrOFSnapshots) = .timeStamp.ToString("d")
        End With


        With CType(appInstance.Worksheets(arrWsNames(3)), Excel.Worksheet)
            anzDiagrams = CType(.ChartObjects, Excel.ChartObjects).Count
            '
            ' um welches Diagramm handelt es sich ...
            '
            i = 1
            found = False
            While i <= anzDiagrams And Not found
                Try
                    chtTitle = CType(.ChartObjects(i), Excel.ChartObject).Chart.ChartTitle.Text
                Catch ex As Exception
                    chtTitle = " "
                End Try

                If chtTitle = diagramTitle Then
                    found = True

                Else
                    i = i + 1
                End If

            End While

            If found Then
                Call MsgBox("Chart wird bereits angezeigt ...")
                appInstance.EnableEvents = formerEE
                repObj = CType(.ChartObjects(i), Excel.ChartObject)
                'appInstance.ScreenUpdating = formerSU
                Exit Sub
            Else
                With appInstance.Charts.Add
                    ' remove extra series
                    Do Until .SeriesCollection.Count = 0
                        .SeriesCollection(1).Delete()
                    Loop

                    .HasAxis(Excel.XlAxisType.xlCategory) = True
                    .HasAxis(Excel.XlAxisType.xlValue) = True

                    With .Axes(Excel.XlAxisType.xlCategory)
                        .HasTitle = False
                        .TickLabelPosition = Excel.XlTickLabelPosition.xlTickLabelPositionHigh
                        '.MinimumScale = 0
                        'With .AxisTitle
                        '    .Characters.text = "Monate"
                        '    .Font.Size = 8
                        'End With
                    End With

                    With .Axes(Excel.XlAxisType.xlValue)
                        .HasTitle = False
                        .HasMajorGridlines = False
                        .HasMinorGridlines = False
                        '.MaximumScale = maxscale
                        '.MinimumScale = 0

                        'With .AxisTitle
                        '    .Characters.text = "Kosten"
                        '    .Font.Size = 8
                        'End With
                    End With

                    .HasLegend = True
                    With .Legend
                        .Position = Excel.Constants.xlRight
                        .Font.Size = awinSettings.fontsizeLegend + 6
                    End With
                    .HasTitle = True
                    .ChartTitle.Text = diagramTitle
                    .ChartTitle.Font.Size = awinSettings.fontsizeTitle + 8
                    .Location(Where:=XlChartLocation.xlLocationAsObject, Name:=appInstance.Worksheets(arrWsNames(3)).name)

                End With

                chtobj = CType(.ChartObjects(anzDiagrams + 1), Excel.ChartObject)

            End If


            'Dim sc As Excel.Series
            'With sc
            '    .MarkerStyle = Excel.XlMarkerStyle.xlMarkerStyleSquare
            '    .MarkerSize = 5
            '    .MarkerForegroundColor = XlRgbColor.rgbDarkGreen
            '    .Format.Line.ForeColor.RGB = XlRgbColor.rgbDarkGreen
            'End With

            With chtobj.Chart

                ' hier werden jetzt die ganzen Series aufgebaut 
                ' Erloes
                With .SeriesCollection.NewSeries
                    .ChartType = Excel.XlChartType.xlLine
                    .name = "Budget"
                    .Values = erloes
                    .XValues = Xdatenreihe
                    .MarkerStyle = Excel.XlMarkerStyle.xlMarkerStyleTriangle
                    .MarkerSize = 3
                    .MarkerForegroundColor = XlRgbColor.rgbDarkGreen
                    With .Format.Line
                        .ForeColor.RGB = XlRgbColor.rgbDarkGreen
                        .Weight = 3
                    End With


                End With

                ' Personalkosten
                With .SeriesCollection.NewSeries
                    .ChartType = Excel.XlChartType.xlLine
                    .name = "Personalkosten"
                    '.Interior.color = farbeExterne
                    .Values = personalKosten
                    .XValues = Xdatenreihe
                    .MarkerStyle = Excel.XlMarkerStyle.xlMarkerStyleSquare
                    .MarkerSize = 3
                    .MarkerForegroundColor = XlRgbColor.rgbDarkRed
                    With .Format.Line
                        .ForeColor.RGB = XlRgbColor.rgbDarkRed
                        .Weight = 3
                    End With
                End With

                ' Sonstige Kosten
                With .SeriesCollection.NewSeries
                    .name = "Sonstige Kosten"
                    '.Interior.color = farbeInternOP
                    .Values = sonstKosten
                    .XValues = Xdatenreihe
                    .ChartType = Excel.XlChartType.xlLine
                    .MarkerStyle = Excel.XlMarkerStyle.xlMarkerStyleCircle
                    .MarkerSize = 3
                    .MarkerForegroundColor = XlRgbColor.rgbDarkOrange
                    With .Format.Line
                        .ForeColor.RGB = XlRgbColor.rgbDarkOrange
                        .Weight = 3
                    End With
                End With

                ' Risiko Kosten
                With .SeriesCollection.NewSeries
                    .name = "Risikokosten"
                    '.Interior.color = iProjektFarbe
                    .Values = risikoKosten
                    .XValues = Xdatenreihe
                    .ChartType = Excel.XlChartType.xlLine
                    .MarkerStyle = Excel.XlMarkerStyle.xlMarkerStyleDot
                    .MarkerSize = 3
                    .MarkerForegroundColor = XlRgbColor.rgbLightGray
                    With .Format.Line
                        .ForeColor.RGB = XlRgbColor.rgbLightGray
                        .Weight = 3
                    End With
                End With

                ' estim Profit
                With .SeriesCollection.NewSeries
                    .name = "progn.Ergebnis"
                    '.Interior.color = ergebnisfarbe2
                    .Values = estimProfit
                    .XValues = Xdatenreihe
                    .ChartType = Excel.XlChartType.xlLine
                    .MarkerStyle = Excel.XlMarkerStyle.xlMarkerStyleDiamond
                    .MarkerSize = 5
                    .MarkerForegroundColor = XlRgbColor.rgbDarkGreen
                    With .Format.Line
                        .ForeColor.RGB = XlRgbColor.rgbDarkGreen
                        .Weight = 3
                    End With

                End With

            End With


            With chtobj
                .Top = top
                .Height = height
                .Left = left
                .Width = width
            End With

        End With

        repObj = chtobj

        'Call awinScrollintoView()
        appInstance.EnableEvents = formerEE
        'appInstance.ScreenUpdating = formerSU



    End Sub

    '
    ' 
    '
    ' 

    ''' <summary>
    ''' zeigt die Entwicklung des strategischen Fits / Risikos über die gesamte Projekt-Historie an 
    ''' Voraussetzung: die Projekt-Historie ist bestimmt 
    ''' </summary>
    ''' <param name="repObj">Verweis auf das erzeugte Chart</param>
    ''' <param name="top">Diagramm-Koordinate oben</param>
    ''' <param name="left">Diagramm-Koordinate links</param>
    ''' <param name="height">Diagramm-Höhe</param>
    ''' <param name="width">Diagramm-Breite</param>
    ''' <remarks></remarks>
    Public Sub createTrendSfit(ByRef repObj As Excel.ChartObject, ByVal top As Double, left As Double, height As Double, width As Double)

        Dim diagramTitle As String = " "
        Dim anzDiagrams As Integer
        Dim found As Boolean

        Dim i As Integer
        Dim Xdatenreihe() As String


        Dim chtTitle As String

        Dim chtobj As Excel.ChartObject
        Dim ErgebnisListeR As New Collection

        Dim nrOFSnapshots As Integer = projekthistorie.Count

        If nrOFSnapshots = 0 Then
            Exit Sub
            'Call MsgBox("es gibt keine Historie ...")
        End If

        Dim strategicFit() As Double
        Dim risiko() As Double




        Dim formerEE As Boolean = appInstance.EnableEvents
        'Dim formerSU As Boolean = appInstance.ScreenUpdating
        appInstance.EnableEvents = False
        'appInstance.ScreenUpdating = False

        ReDim strategicFit(nrOFSnapshots)
        ReDim risiko(nrOFSnapshots)
        ReDim Xdatenreihe(nrOFSnapshots)


        Dim pname As String = projekthistorie.Last.name

        diagramTitle = "Planungs-Historie strategischer Fit & Risiko: " & vbLf & projekthistorie.Last.getShapeText


        ' jetzt werden die einzelnen Werte aufgefüllt
        Dim ix As Integer = 0
        For Each kvp As KeyValuePair(Of Date, clsProjekt) In projekthistorie.liste

            With kvp.Value
                strategicFit(ix) = .StrategicFit
                risiko(ix) = .Risiko
                Xdatenreihe(ix) = .timeStamp.ToString("d")
            End With

            ix = ix + 1

        Next

        Dim hproj As clsProjekt = ShowProjekte.getProject(pname)
        With hproj

            strategicFit(nrOFSnapshots) = .StrategicFit
            risiko(nrOFSnapshots) = .Risiko
            Xdatenreihe(nrOFSnapshots) = .timeStamp.ToString("d")

        End With


        With CType(appInstance.Worksheets(arrWsNames(3)), Excel.Worksheet)
            anzDiagrams = CType(.ChartObjects, Excel.ChartObjects).Count
            '
            ' um welches Diagramm handelt es sich ...
            '
            i = 1
            found = False
            While i <= anzDiagrams And Not found
                Try
                    chtTitle = CType(.ChartObjects(i), Excel.ChartObject).Chart.ChartTitle.Text
                Catch ex As Exception
                    chtTitle = " "
                End Try

                If chtTitle = diagramTitle Then
                    found = True

                Else
                    i = i + 1
                End If

            End While

            If found Then
                Call MsgBox("Chart wird bereits angezeigt ...")
                appInstance.EnableEvents = formerEE
                'appInstance.ScreenUpdating = formerSU
                repObj = CType(.ChartObjects(i), Excel.ChartObject)
                Exit Sub
            Else
                With appInstance.Charts.Add
                    ' remove extra series
                    Do Until .SeriesCollection.Count = 0
                        .SeriesCollection(1).Delete()
                    Loop

                    .HasAxis(Excel.XlAxisType.xlCategory) = True
                    .HasAxis(Excel.XlAxisType.xlValue) = True

                    With .Axes(Excel.XlAxisType.xlCategory)
                        .HasTitle = False
                        .TickLabelPosition = Excel.XlTickLabelPosition.xlTickLabelPositionHigh
                        '.MinimumScale = 0
                        'With .AxisTitle
                        '    .Characters.text = "Monate"
                        '    .Font.Size = 8
                        'End With
                    End With


                    'Dim ax As Excel.Axis
                    'With ax
                    '    .HasMajorGridlines = False
                    '    .HasMinorGridlines = False
                    'End With
                    With .Axes(Excel.XlAxisType.xlValue)
                        .HasTitle = False
                        .HasMajorGridlines = False
                        .HasMinorGridlines = False
                        .MaximumScale = 11
                        .MinimumScale = 0

                        'With .AxisTitle
                        '    .Characters.text = "Kosten"
                        '    .Font.Size = 8
                        'End With
                    End With

                    .HasLegend = True
                    With .Legend
                        .Position = Excel.Constants.xlRight
                        .Font.Size = awinSettings.fontsizeLegend + 6
                    End With
                    .HasTitle = True
                    .ChartTitle.Text = diagramTitle
                    .ChartTitle.Font.Size = awinSettings.fontsizeTitle + 8
                    .Location(Where:=XlChartLocation.xlLocationAsObject, Name:=appInstance.Worksheets(arrWsNames(3)).name)

                End With

                chtobj = CType(.ChartObjects(anzDiagrams + 1), Excel.ChartObject)

            End If


            'Dim sc As Excel.Series
            'With sc
            '    .MarkerStyle = Excel.XlMarkerStyle.xlMarkerStyleSquare
            '    .MarkerSize = 5
            '    .MarkerForegroundColor = XlRgbColor.rgbDarkGreen
            '    .Format.Line.ForeColor.RGB = XlRgbColor.rgbDarkGreen
            'End With

            With chtobj.Chart

                ' hier werden jetzt die ganzen Series aufgebaut 
                ' Erloes
                With .SeriesCollection.NewSeries
                    .ChartType = Excel.XlChartType.xlLine
                    .name = "Strategischer Fit"
                    .Values = strategicFit
                    .XValues = Xdatenreihe
                    .MarkerStyle = Excel.XlMarkerStyle.xlMarkerStyleStar
                    .MarkerSize = 5
                    .MarkerForegroundColor = XlRgbColor.rgbDarkGreen
                    With .Format.Line
                        .ForeColor.RGB = XlRgbColor.rgbDarkGreen
                        .Weight = 3
                    End With


                End With

                ' Personalkosten
                With .SeriesCollection.NewSeries
                    .ChartType = Excel.XlChartType.xlLine
                    .name = "Risiko"
                    '.Interior.color = farbeExterne
                    .Values = risiko
                    .XValues = Xdatenreihe
                    .MarkerStyle = Excel.XlMarkerStyle.xlMarkerStyleSquare
                    .MarkerSize = 5
                    .MarkerForegroundColor = XlRgbColor.rgbDarkOrange
                    With .Format.Line
                        .ForeColor.RGB = XlRgbColor.rgbDarkOrange
                        .Weight = 3
                    End With
                End With


            End With


            With chtobj
                .Top = top
                .Height = height
                .Left = left
                .Width = width
            End With

        End With

        repObj = chtobj

        'Call awinScrollintoView()
        appInstance.EnableEvents = formerEE
        'appInstance.ScreenUpdating = formerSU



    End Sub


    ''' <summary>
    ''' zeigt die Charakteristik aufgelöst nach den einzelnen Monaten 
    ''' wird momentan nicht benutzt !
    ''' </summary>
    ''' <param name="pname">Name des Projekts</param>
    ''' <param name="auswahl"></param>
    ''' <remarks></remarks>
    Public Sub createProjektErgebnisCharakteristik(ByVal pname As String, auswahl As Integer)

        Dim diagramTitle As String
        Dim anzDiagrams As Integer
        Dim found As Boolean
        Dim plen As Integer
        Dim i As Integer
        Dim minScale As Double
        Dim Xdatenreihe() As String
        Dim earnedValues() As Double
        Dim earnedValuesWeighted() As Double
        Dim riskValues() As Double
        Dim hsum(1) As Double, gesamtSumme As Double
        Dim top As Double, left As Double, width As Double, height As Double
        Dim chtTitle As String
        Dim pstart As Integer
        Dim mycollection As New Collection

        Dim hproj As clsProjekt
        Dim ErgebnisListeR As New Collection


        appInstance.EnableEvents = False
        appInstance.ScreenUpdating = False

        If auswahl = 1 Then
            diagramTitle = "Projekt Earned Value"
        ElseIf auswahl = 2 Then
            diagramTitle = "Projekt Earned Value mit Risiko-Abschlag"
        Else
            diagramTitle = pname
        End If

        hproj = ShowProjekte.getProject(pname)
        '
        ' hole die Projektdauer
        '
        plen = hproj.anzahlRasterElemente
        pstart = hproj.Start


        ReDim Xdatenreihe(plen - 1)
        ReDim earnedValues(plen - 1)
        ReDim earnedValuesWeighted(plen - 1)
        ReDim riskValues(plen - 1)



        For i = 1 To plen
            Xdatenreihe(i - 1) = StartofCalendar.AddMonths(pstart + i - 2).ToString("MMM yy")
        Next i


        ' neu 
        '
        ' die Position des Diagramms wird ausgerechnet ...
        '
        top = 48 + (hproj.tfZeile - 1) * 15
        left = (hproj.tfspalte - 1) * boxWidth - 5
        If left < 0 Then
            left = 0
        End If
        height = 180
        width = plen * boxWidth + 10


        gesamtSumme = 0

        With CType(appInstance.Worksheets(arrWsNames(3)), Excel.Worksheet)
            anzDiagrams = CType(.ChartObjects, Excel.ChartObjects).Count
            '
            ' um welches Diagramm handelt es sich ...
            '
            i = 1
            found = False
            While i <= anzDiagrams And Not found
                Try
                    chtTitle = CType(.ChartObjects(i), Excel.ChartObject).Chart.ChartTitle.Text
                Catch ex As Exception
                    chtTitle = " "
                End Try
                If chtTitle = diagramTitle Then
                    found = True
                Else
                    i = i + 1
                End If

            End While

            If found Then
                MsgBox(" Diagramm wird bereits angezeigt ...")
            Else

                With appInstance.Charts.Add
                    ' remove extra series
                    Do Until .SeriesCollection.Count = 0
                        .SeriesCollection(1).Delete()
                    Loop

                    mycollection.Add(ergebnisChartName(0))
                    earnedValues = hproj.getBedarfeInMonths(mycollection, DiagrammTypen(4))
                    mycollection.Clear()

                    If auswahl = 1 Then
                        'minScale = earnedValues.Min
                        'mycollection.Clear()
                    ElseIf auswahl = 2 Then ' es sind die um die Risiko-Abschläge bereinigten Earned Values
                        mycollection.Add(ergebnisChartName(1))
                        earnedValues = hproj.getBedarfeInMonths(mycollection, DiagrammTypen(4))
                        mycollection.Clear()

                        mycollection.Add(ergebnisChartName(3))
                        riskValues = hproj.getBedarfeInMonths(mycollection, DiagrammTypen(4))
                        mycollection.Clear()


                    End If

                    minScale = earnedValues.Min

                    hsum(0) = 0
                    For i = 0 To plen - 1
                        hsum(0) = hsum(0) + earnedValues(i)
                    Next i
                    gesamtSumme = hsum(0)

                    If auswahl = 2 Then
                        hsum(1) = 0
                        For i = 0 To plen - 1
                            hsum(1) = hsum(1) + riskValues(i)
                        Next i
                        gesamtSumme = gesamtSumme + hsum(1)
                    End If


                    'series
                    With .SeriesCollection.NewSeries
                        If auswahl = 1 Then
                            .name = ergebnisChartName(0)
                        Else
                            .name = ergebnisChartName(1)
                        End If
                        .Interior.color = ergebnisfarbe1
                        .Values = earnedValues
                        .XValues = Xdatenreihe
                        .ChartType = Excel.XlChartType.xlColumnStacked
                    End With

                    'If auswahl = 2 Then
                    '    With .SeriesCollection.NewSeries
                    '        .name = "Risiko Abschlag"
                    '        .Interior.color = ergebnisfarbe2
                    '        .Values = riskValues
                    '        .XValues = Xdatenreihe
                    '        .ChartType = Excel.XlChartType.xlColumnStacked
                    '    End With
                    'End If


                    .HasAxis(Excel.XlAxisType.xlCategory) = True
                    .HasAxis(Excel.XlAxisType.xlValue) = True

                    With .Axes(Excel.XlAxisType.xlCategory)
                        .HasTitle = False
                        If minScale < 0 Then
                            .TickLabelPosition = Excel.Constants.xlLow
                        End If
                        '.MinimumScale = 0
                        'With .AxisTitle
                        '    .Characters.text = "Monate"
                        '    .Font.Size = 8
                        'End With
                    End With

                    With .Axes(Excel.XlAxisType.xlValue)
                        .HasTitle = False
                        If minScale < 0 Then
                            .MinimumScale = System.Math.Round(minScale - 1, mode:=MidpointRounding.ToEven)
                            '.AxisTitle.Position = Excel.XlChartElementPosition.xlChartElementPositionCustom
                            '.AxisTitle.Position = XlConstants.xlBottom
                        Else
                            .MinimumScale = 0
                        End If

                        'Dim hax As Excel.Axis
                        'With hax
                        '    .AxisTitle.Position = Excel.XlChartElementPosition.xlChartElementPositionCustom
                        'End With

                        'With .AxisTitle
                        '    If auswahl = 1 Then
                        '        .Characters.text = "Ressourcen"
                        '    Else
                        '        .Characters.text = "Personalkosten"
                        '    End If
                        '    .Font.Size = 8
                        'End With
                    End With

                    'If auswahl = 2 Then
                    '    .HasLegend = True
                    '    With .Legend
                    '        .Position = XlConstants.xlTop
                    '        .Font.Size = 8
                    '    End With
                    'Else
                    '    .HasLegend = False
                    'End If
                    .HasLegend = False
                    .HasTitle = True
                    diagramTitle = diagramTitle & " " & Format(hsum(0), "##,##0") & " T€" & vbLf & pname
                    .ChartTitle.Text = diagramTitle
                    .ChartTitle.font.size = 10
                    .Location(Where:=XlChartLocation.xlLocationAsObject, Name:=appInstance.Worksheets(arrWsNames(3)).name)
                End With

                ' jetzt kommt die Korrektur der Größe; herausfinden, wieviel Raum die Axis Beschriftung einnimmt ... 
                With .ChartObjects(anzDiagrams + 1)
                    .top = top
                    .height = 2 * height

                    Dim axleft As Double, axwidth As Double
                    If .Chart.HasAxis(Excel.XlAxisType.xlValue) = True Then
                        With CType(.Chart.Axes(Excel.XlAxisType.xlValue), Excel.Axis)
                            axleft = .Left
                            axwidth = .Width
                        End With
                        If left - axwidth < 1 Then
                            left = 1
                            width = width + left + 9
                        Else
                            left = left - axwidth
                            width = width + axwidth + 9
                        End If

                    End If

                    .left = left
                    .width = width


                End With



            End If


        End With


        appInstance.EnableEvents = True
        appInstance.ScreenUpdating = True


    End Sub

    ''' <summary>
    ''' diese Sub zeigt das vorauss. Projektergebnis in einem Chart an - Erloes, Risiko Kosten, Personalkosten, Sonstige Kosten und 
    ''' dann den vermutl. Projekt-Ertrag
    ''' auswahl: 0=beauftragung; 1: letzter Stand; 2: aktueller Stand   
    ''' </summary>
    ''' <param name="hproj">das Projekt</param>
    ''' <param name="reportObj">
    ''' nimmt den Verweis auf das generierte Chart auf; 
    ''' wird für das Reporting benötigt 
    ''' </param>
    ''' <remarks></remarks>
    Public Sub createProjektErgebnisCharakteristik2(ByRef hproj As clsProjekt, ByRef reportObj As Excel.ChartObject, ByVal auswahl As Integer)

        Dim diagramTitle As String
        Dim anzDiagrams As Integer
        Dim found As Boolean
        Dim plen As Integer
        Dim i As Integer
        Dim minScale As Double
        Dim Xdatenreihe(4) As String
        Dim valueDatenreihe1(4) As Double
        Dim valueDatenreihe2(4) As Double
        Dim itemColor(4) As Object
        Dim itemValue(4) As Double
        Dim projektErloes As Double, projektPersKosten As Double, projektSonstKosten As Double, projektRisikoKosten As Double
        Dim projektErgebnis As Double
        'Dim earnedValueWeighted As Double
        Dim top As Double, left As Double, width As Double, height As Double
        Dim pstart As Integer
        Dim mycollection As New Collection
        'Dim catName As String
        Dim pname As String = hproj.name
        Dim titelTeile(1) As String
        Dim titelTeilLaengen(1) As Integer
        Dim kennung As String
        Dim tmpcollection As New Collection


        tmpcollection.Add(hproj.getShapeText & "#" & auswahl.ToString)
        kennung = calcChartKennung("pr", PTprdk.Ergebnis, tmpcollection)


        '
        ' hole die Projektdauer
        '
        With hproj
            plen = .anzahlRasterElemente
            pstart = .Start
        End With


        Xdatenreihe(0) = "Budget"
        Xdatenreihe(1) = "Risiko-Abschlag"
        Xdatenreihe(2) = "Personalkosten"
        Xdatenreihe(3) = "Sonstige Kosten"
        Xdatenreihe(4) = "Ergebnis-Prognose"



        With hproj

            .calculateRoundedKPI(projektErloes, projektPersKosten, projektSonstKosten, projektRisikoKosten, projektErgebnis)

            itemValue(0) = projektErloes
            itemColor(0) = ergebnisfarbe1

            itemValue(1) = projektRisikoKosten
            itemColor(1) = iProjektFarbe

            itemValue(2) = projektPersKosten
            itemColor(2) = farbeExterne

            itemValue(3) = projektSonstKosten
            itemColor(3) = farbeInternOP

            itemValue(4) = projektErgebnis
            If projektErgebnis > 0 Then
                itemColor(4) = ergebnisfarbe2
            Else
                itemColor(4) = farbeExterne
            End If
        End With


        If auswahl = PThis.beauftragung Then
            titelTeile(0) = hproj.getShapeText & " (Beauftragung)" & vbLf & textZeitraum(pstart, pstart + plen - 1) & vbLf
        ElseIf auswahl = PThis.letzterStand Then
            titelTeile(0) = hproj.getShapeText & " (letzter Stand)" & vbLf & textZeitraum(pstart, pstart + plen - 1) & vbLf
        Else
            titelTeile(0) = hproj.getShapeText & vbLf & textZeitraum(pstart, pstart + plen - 1) & vbLf
        End If

        titelTeilLaengen(0) = titelTeile(0).Length
        titelTeile(1) = " (" & hproj.timeStamp.ToString & ") "
        titelTeilLaengen(1) = titelTeile(1).Length

        diagramTitle = titelTeile(0) & titelTeile(1)
        'kennung = pname & "#Ergebnis#1"


        ' neu 
        '
        ' die Position des Diagramms wird ausgerechnet ...
        '
        top = topOfMagicBoard + hproj.tfZeile * boxHeight
        left = hproj.tfspalte * boxWidth - 10
        If left < 0 Then
            left = 1
        End If
        height = awinSettings.ChartHoehe2
        width = 450






        With CType(appInstance.Worksheets(arrWsNames(3)), Excel.Worksheet)
            anzDiagrams = CType(.ChartObjects, Excel.ChartObjects).Count

            '
            ' um welches Diagramm handelt es sich ...
            '
            i = 1
            found = False
            While i <= anzDiagrams And Not found


                If kennung = .ChartObjects(i).name Then
                    found = True
                Else
                    i = i + 1
                End If

            End While


            Dim currentWert As Double
            If found Then
                reportObj = CType(.ChartObjects(i), Excel.ChartObject)
            Else

                If projektErgebnis < 0 Then
                    minScale = System.Math.Round(projektErgebnis, mode:=MidpointRounding.ToEven)
                Else
                    minScale = 0
                End If


                Dim valueCrossesNull As Boolean = False

                With appInstance.Charts.Add
                    ' remove extra series
                    Do Until .SeriesCollection.Count = 0
                        .SeriesCollection(1).Delete()
                    Loop
                    Dim crossindex As Integer = -1

                    ' bestimmen des Anfangs  
                    Dim iv = 0
                    valueDatenreihe1(iv) = 0
                    valueDatenreihe2(iv) = itemValue(iv)
                    currentWert = itemValue(iv)

                    ' alle nächsten Zwischen-Werte 
                    Dim negativeFromNull As Boolean = False
                    Dim formerValue As Double = currentWert
                    For iv = 1 To 3
                        If formerValue <= 0 Then
                            negativeFromNull = True
                        Else
                            negativeFromNull = False
                        End If

                        currentWert = currentWert - itemValue(iv)
                        valueCrossesNull = (currentWert + itemValue(iv) > 0) And (currentWert < 0)

                        If currentWert >= 0 Then
                            valueDatenreihe1(iv) = currentWert
                            valueDatenreihe2(iv) = itemValue(iv)
                        ElseIf valueCrossesNull Then
                            valueDatenreihe1(iv) = currentWert
                            valueDatenreihe2(iv) = itemValue(iv) - currentWert * (-1) ' notwendig da currentWert ja negativ ist ..
                            crossindex = iv + 1
                        ElseIf negativeFromNull Then
                            valueDatenreihe1(iv) = formerValue
                            valueDatenreihe2(iv) = itemValue(iv) * (-1)
                        Else
                            valueDatenreihe1(iv) = currentWert
                            valueDatenreihe2(iv) = itemValue(iv) * (-1)
                        End If

                        formerValue = currentWert
                    Next

                    ' bestimmen des Ende 
                    iv = 4
                    valueDatenreihe1(iv) = 0
                    valueDatenreihe2(iv) = itemValue(iv)



                    'series
                    With .SeriesCollection.NewSeries
                        .name = "Bottom"
                        .HasDataLabels = False
                        .Interior.colorindex = -4142
                        .Values = valueDatenreihe1
                        .XValues = Xdatenreihe
                        .ChartType = Excel.XlChartType.xlColumnStacked
                        If crossindex > 0 Then
                            ' es gab einen Übergang , dort muss Bottom auf die entsprechende Farbe gesetzt werden 
                            With .Points(crossindex)
                                .Interior.color = itemColor(crossindex - 1)
                            End With
                        End If

                    End With

                    With .SeriesCollection.NewSeries
                        .name = "Top"
                        .HasDataLabels = True
                        .Values = valueDatenreihe2
                        .XValues = Xdatenreihe
                        .ChartType = Excel.XlChartType.xlColumnStacked

                        For iv = 0 To 4

                            With .Points(iv + 1)
                                .HasDataLabel = True
                                .DataLabel.text = Format(itemValue(iv), "###,###0") & " T€"
                                .Interior.color = itemColor(iv)
                                .DataLabel.Font.Size = awinSettings.fontsizeItems + 2
                                'Try
                                '    .DataLabel.Position = Excel.XlDataLabelPosition.xlLabelPositionAbove
                                'Catch ex As Exception

                                'End Try
                            End With

                        Next

                    End With

                    .HasAxis(Excel.XlAxisType.xlCategory) = True
                    .HasAxis(Excel.XlAxisType.xlValue) = False



                    With .Axes(Excel.XlAxisType.xlCategory)
                        .HasTitle = False
                        If minScale < 0 Then
                            .TickLabelPosition = Excel.Constants.xlLow
                        End If
                        '.MinimumScale = 0

                    End With

                    'Dim hax As Excel.Axis
                    'With hax
                    '    .HasMajorGridlines
                    '    .hasminor()
                    'End With

                    Try
                        With .Axes(Excel.XlAxisType.xlValue)
                            .HasTitle = False
                            .HasMajorGridlines = False
                            .hasminorgridlines = False
                            If minScale < 0 Then
                                .MinimumScale = System.Math.Round((minScale - 1), mode:=MidpointRounding.ToEven)
                            Else
                                .MinimumScale = 0
                            End If
                        End With
                    Catch ex As Exception

                    End Try


                    .HasLegend = False
                    'With .Legend
                    '    .Position = XlConstants.xlTop
                    '    .Font.Size = 8
                    'End With
                    .HasTitle = True

                    .ChartTitle.Text = diagramTitle
                    .ChartTitle.Font.Size = awinSettings.fontsizeTitle
                    .ChartTitle.Format.TextFrame2.TextRange.Characters(titelTeilLaengen(0) + 1, _
                        titelTeilLaengen(1)).Font.Size = awinSettings.fontsizeLegend

                    .Location(Where:=XlChartLocation.xlLocationAsObject, Name:=appInstance.Worksheets(arrWsNames(3)).name)
                End With



                With .ChartObjects(anzDiagrams + 1)
                    .top = top
                    .left = left
                    .width = width
                    .height = height
                    .name = kennung

                End With

                reportObj = CType(.ChartObjects(anzDiagrams + 1), Excel.ChartObject)


            End If


        End With





    End Sub

    Public Sub updateProjektErgebnisCharakteristik2(ByRef hproj As clsProjekt, ByRef chtobj As Excel.ChartObject, _
                                                    ByVal auswahl As Integer, ByVal changeScale As Boolean)


        Dim diagramTitle As String

        Dim plen As Integer
        Dim Xdatenreihe(4) As String
        Dim valueDatenreihe1(4) As Double
        Dim valueDatenreihe2(4) As Double
        Dim itemColor(4) As Object
        Dim itemValue(4) As Double
        Dim projektErloes As Double, projektPersKosten As Double, projektSonstKosten As Double, projektRisikoKosten As Double
        Dim projektErgebnis As Double
        'Dim earnedValueWeighted As Double

        Dim pstart As Integer
        Dim mycollection As New Collection
        'Dim catName As String
        Dim pname As String = hproj.name
        Dim minscale As Double

        Dim titelTeile(1) As String
        Dim titelTeilLaengen(1) As Integer
        Dim kennung As String
        Dim tmpcollection As New Collection

        Dim formerEE As Boolean = appInstance.EnableEvents
        appInstance.EnableEvents = False

        tmpcollection.Add(hproj.name & "#" & auswahl.ToString)
        kennung = calcChartKennung("pr", PTprdk.Ergebnis, tmpcollection)


        '
        ' hole die Projektdauer
        '
        With hproj
            plen = .anzahlRasterElemente
            pstart = .Start
        End With


        Xdatenreihe(0) = "Budget"
        Xdatenreihe(1) = "Risiko-Abschlag"
        Xdatenreihe(2) = "Personalkosten"
        Xdatenreihe(3) = "Sonstige Kosten"
        Xdatenreihe(4) = "Ergebnis-Prognose"



        With hproj
            Dim gk As Double = .getSummeKosten
            projektErloes = System.Math.Round(.Erloes, mode:=MidpointRounding.ToEven)
            itemValue(0) = projektErloes
            itemColor(0) = ergebnisfarbe1

            projektRisikoKosten = System.Math.Round(.risikoKostenfaktor * gk, mode:=MidpointRounding.ToEven)
            itemValue(1) = projektRisikoKosten
            itemColor(1) = iProjektFarbe

            projektPersKosten = System.Math.Round(.getAllPersonalKosten.Sum, mode:=MidpointRounding.ToEven)
            itemValue(2) = projektPersKosten
            itemColor(2) = farbeExterne

            projektSonstKosten = System.Math.Round(.getGesamtAndereKosten.Sum, mode:=MidpointRounding.ToEven)
            itemValue(3) = projektSonstKosten
            itemColor(3) = farbeInternOP

            projektErgebnis = projektErloes - (projektRisikoKosten + projektPersKosten + projektSonstKosten)
            itemValue(4) = projektErgebnis
            If projektErgebnis > 0 Then
                itemColor(4) = ergebnisfarbe2
            Else
                itemColor(4) = farbeExterne
            End If
        End With



        titelTeile(0) = pname & vbLf & textZeitraum(pstart, pstart + plen - 1) & vbLf
        titelTeilLaengen(0) = titelTeile(0).Length
        titelTeile(1) = " (" & hproj.timeStamp.ToString & ") "
        titelTeilLaengen(1) = titelTeile(1).Length
        diagramTitle = titelTeile(0) & titelTeile(1)
        'kennung = pname & "#Ergebnis#1"


        If changeScale Then
            If projektErgebnis < 0 Then
                minscale = System.Math.Round(projektErgebnis - 5, mode:=MidpointRounding.ToEven)

                If projektErgebnis < -300 Then
                    minscale = Math.Round(projektErgebnis / 50 - 0.6) * 50
                ElseIf projektErgebnis < -80 Then
                    minscale = Math.Round(projektErgebnis / 10 - 0.6) * 10
                Else
                    minscale = Math.Round(projektErgebnis / 5 - 0.6) * 5
                End If


            Else
                minscale = 0
            End If
        End If
        


        Dim currentWert As Double



        Dim valueCrossesNull As Boolean = False

        With chtobj.Chart

            'ur:22.07.2014: bereits in Charts enthalten und soll nur mit neuen Daten bestückt werden
            '' remove extra series

            'Do Until .SeriesCollection.Count = 0
            '    .SeriesCollection(1).Delete()
            'Loop

            Dim crossindex As Integer = -1

            ' bestimmen des Anfangs  
            Dim iv = 0
            valueDatenreihe1(iv) = 0
            valueDatenreihe2(iv) = itemValue(iv)
            currentWert = itemValue(iv)

            ' alle nächsten Zwischen-Werte 
            Dim negativeFromNull As Boolean = False
            Dim formerValue As Double = currentWert
            For iv = 1 To 3
                If formerValue <= 0 Then
                    negativeFromNull = True
                Else
                    negativeFromNull = False
                End If

                currentWert = currentWert - itemValue(iv)
                valueCrossesNull = (currentWert + itemValue(iv) > 0) And (currentWert < 0)

                If currentWert >= 0 Then
                    valueDatenreihe1(iv) = currentWert
                    valueDatenreihe2(iv) = itemValue(iv)
                ElseIf valueCrossesNull Then
                    valueDatenreihe1(iv) = currentWert
                    valueDatenreihe2(iv) = itemValue(iv) - currentWert * (-1) ' notwendig da currentWert ja negativ ist ..
                    crossindex = iv + 1
                ElseIf negativeFromNull Then
                    valueDatenreihe1(iv) = formerValue
                    valueDatenreihe2(iv) = itemValue(iv) * (-1)
                Else
                    valueDatenreihe1(iv) = currentWert
                    valueDatenreihe2(iv) = itemValue(iv) * (-1)
                End If

                formerValue = currentWert
            Next

            ' bestimmen des Ende 
            iv = 4
            valueDatenreihe1(iv) = 0
            valueDatenreihe2(iv) = itemValue(iv)



            'series
            'With .SeriesCollection.NewSeries
            With .SeriesCollection(1)
                .name = "Bottom"
                .HasDataLabels = False
                .Interior.colorindex = -4142
                .Values = valueDatenreihe1
                .XValues = Xdatenreihe
                .ChartType = Excel.XlChartType.xlColumnStacked
                If crossindex > 0 Then
                    ' es gab einen Übergang , dort muss Bottom auf die entsprechende Farbe gesetzt werden 
                    With .Points(crossindex)
                        .Interior.color = itemColor(crossindex - 1)
                    End With
                End If

            End With


            'With .SeriesCollection.NewSeries
            With .SeriesCollection(2)
                .name = "Top"
                .HasDataLabels = True
                .Values = valueDatenreihe2
                .XValues = Xdatenreihe
                .ChartType = Excel.XlChartType.xlColumnStacked

                For iv = 0 To 4

                    With .Points(iv + 1)
                        .HasDataLabel = True
                        .DataLabel.text = Format(itemValue(iv), "###,###0") & " T€"
                        .Interior.color = itemColor(iv)
                        '.DataLabel.Font.Size = awinSettings.fontsizeItems + 2
                        'Try
                        '    .DataLabel.Position = Excel.XlDataLabelPosition.xlLabelPositionAbove
                        'Catch ex As Exception

                        'End Try
                    End With

                Next

            End With

            With .Axes(Excel.XlAxisType.xlCategory)
                .HasTitle = False
                If minscale < 0 Then
                    .TickLabelPosition = Excel.Constants.xlLow
                End If


            End With


            Try
                With .Axes(Excel.XlAxisType.xlValue)
                    .HasTitle = False
                    .HasMajorGridlines = False
                    .hasminorgridlines = False
                End With
            Catch ex As Exception

            End Try


            With CType(.Axes(Excel.XlAxisType.xlValue), Excel.Axis)
                ' das ist dann relevant, wenn ein anderes Projekt selektiert wird, das über die aktuelle Skalierung 
                ' hinausgehende Werte hat 
                'If Not (.MaximumScaleIsAuto) Then


                If changeScale Then

                    If Not (.MinimumScaleIsAuto) Then
                        If (minscale < .MinimumScale) And (minscale < 0) Then
                            .MinimumScale = minscale
                        End If
                        .MinimumScaleIsAuto = True
                    End If


                    If Not (.MaximumScaleIsAuto) Then

                        'If itemValue(0) > .MaximumScale - 3 Then
                        '    .MaximumScale = itemValue(0) + 3
                        'End If

                        If itemValue(0) > .MaximumScale Then
                            If itemValue(0) < 80 Then
                                .MaximumScale = Math.Round(itemValue(0) / 5 + 0.6) * 5
                            ElseIf itemValue(0) < 300 Then
                                .MaximumScale = Math.Round(itemValue(0) / 10 + 0.6) * 10
                            Else
                                .MaximumScale = Math.Round(itemValue(0) / 50 + 0.6) * 50
                            End If
                        End If


                        .MaximumScaleIsAuto = True
                    End If
                End If

            End With


            If .HasTitle Then
                .ChartTitle.Text = diagramTitle
                '.ChartTitle.Font.Size = awinSettings.fontsizeTitle
                '.ChartTitle.Format.TextFrame2.TextRange.Characters(titelTeilLaengen(0) + 1, _
                '    titelTeilLaengen(1)).Font.Size = awinSettings.fontsizeLegend
            End If


        End With

        chtobj.Name = kennung

        appInstance.EnableEvents = formerEE

    End Sub



    '
    ' Sub trägt ein individuelles Projekt ein
    '
    ''' <summary>
    ''' trägt ein neues Projekt in Showprojekte ein
    ''' </summary>
    ''' <param name="pname">Projektname</param>
    ''' <param name="vorlagenName">Vorlagen-Name</param>
    ''' <param name="startdate">Start-Datum des PRojekts</param>
    ''' <param name="erloes">Budget des Projekts</param>
    ''' <param name="tafelZeile">
    ''' in welcher Zeile der Projekt-Tafel soll es gezeichnet werden; 
    ''' 0:= finde eine geeignete Stelle
    ''' </param>
    ''' <param name="sfit">Wert für den strategischen Fit</param>
    ''' <param name="risk">Wert für das Risiko</param>
    ''' <param name="volume">Wert für das Volumen</param>
    ''' <remarks></remarks>
    Public Sub TrageivProjektein(ByVal pname As String, ByVal vorlagenName As String, ByVal startdate As Date, _
                                 ByVal endedate As Date, ByVal erloes As Double, _
                                 ByVal tafelZeile As Integer, ByVal sfit As Double, ByVal risk As Double, ByVal volume As Double)
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


        newprojekt = True

        '
        ' ein neues Projekt wird als Objekt angelegt ....
        '

        hproj = New clsProjekt

        Try
            ' Projektdauer wurde durch Start- und Endedatum im Formular angegeben
            Projektvorlagen.getProject(vorlagenName).korrCopyTo(hproj, startdate, endedate)

        Catch ex As Exception
            Call MsgBox("es gibt keine entsprechende Vorlage ..")
            Exit Sub
        End Try


        Try
            With hproj
                .name = pname
                .VorlagenName = vorlagenName
                .startDate = startdate
                .Erloes = erloes
                .earliestStartDate = .startDate.AddMonths(.earliestStart)
                .latestStartDate = .startDate.AddMonths(.latestStart)
                .Status = ProjektStatus(0)

                .volume = volume
                .StrategicFit = sfit
                .Risiko = risk
                plen = .anzahlRasterElemente
                pcolor = .farbe
            End With


            ' nächste Zeile ist ein work-around für Fehler Der Index liegt außerhalb der Array-Grenzen
            ' workaround
            Dim tmpdata As Integer = hproj.dauerInDays
            Call awinCreateBudgetWerte(hproj)

        Catch ex As Exception
            Call MsgBox(ex.Message)
            Exit Sub
        End Try

        ' Anpassen der Daten für die Termine 
        ' wenn Samstag oder Sonntag, dann auf den Freitag davor legen 
        ' nein - das darf nicht gemacht werden; evtl liegt ja dann der Meilenstein vor der Phase 
        ' grundsätzlich sollte der Anwender hier bestimmen, nicht das Programm


        '
        ' Ende Objekt Anlage
        '

        Dim formerEE As Boolean = appInstance.EnableEvents
        Dim formerSU As Boolean = appInstance.ScreenUpdating
        appInstance.EnableEvents = False
        appInstance.ScreenUpdating = False


        Try
            AlleProjekte.Add(key, hproj)
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
        Call ZeichneProjektinPlanTafel(tmpCollection, pname, 0, tmpCollection, tmpCollection)


        '
        ' wenn Röntgen-Blick ein ist, dann müssen die Werte für dieses Projekt eingetragen  werden
        '
        If roentgenBlick.isOn Then
            With roentgenBlick
                Call awinShowNeedsofProject1(mycollection:=.myCollection, type:=.type, projektname:=pname)
            End With
        End If


        ' ein Projekt wurde eingefügt  - typus = 2
        Call awinNeuZeichnenDiagramme(2)

        ' Call diagramsVisible("True")
        appInstance.ScreenUpdating = formerSU
        appInstance.EnableEvents = formerEE

    End Sub


    '
    ' Sub trägt ein individuelles Projekt ein
    '
    Public Sub erstelleInventurProjekt(ByRef hproj As clsProjekt, ByVal pname As String, ByVal vorlagenName As String, ByVal variantName As String, _
                                       ByVal startdate As Date, ByVal endedate As Date, _
                                       ByVal erloes As Double, ByVal tafelZeile As Integer, ByVal sfit As Double, ByVal risk As Double, _
                                       ByVal volume As Double, ByVal complexity As Double, ByVal businessUnit As String, ByVal description As String)

        Dim newprojekt As Boolean
        Dim pStatus As String = ProjektStatus(1) ' jedes Projekt soll zu Beginn als beauftragtes Projekt importiert werden 
        Dim zeile As Integer = tafelZeile
        Dim spalte As Integer = getColumnOfDate(startdate)
        Dim heute As Date = Now
        Dim key As String = pname & "#"

        newprojekt = True

        '
        ' ein neues Projekt wird als Objekt angelegt ....
        '
        If IsNothing(variantName) Then
            variantName = ""
        End If

        Try
            Projektvorlagen.getProject(vorlagenName).korrCopyTo(hproj, startdate, endedate)
            'Projektvorlagen.getProject(vorlagenName).CopyTo(hproj)
        Catch ex As Exception
            Call MsgBox("es gibt keine entsprechende Vorlage ..")
            Exit Sub
        End Try

        Try
            With hproj
                .name = pname
                .variantName = variantName
                '.getPhase(1).name = pname
                .getPhase(1).nameID = rootPhaseName
                .VorlagenName = vorlagenName
                .startDate = startdate
                .earliestStartDate = .startDate.AddMonths(.earliestStart)
                .latestStartDate = .startDate.AddMonths(.latestStart)
                ' jedes Projekt zu Beginn als beauftragtes Projekt importieren
                .Status = ProjektStatus(0)
                .StrategicFit = sfit
                .Risiko = risk

                If Not IsNothing(volume) Then
                    .volume = volume
                Else
                    .volume = 0.0
                End If

                If Not IsNothing(complexity) Then
                    .complexity = complexity
                Else
                    .complexity = 0.0
                End If

                .businessUnit = businessUnit
                .description = description
                .tfZeile = tafelZeile
                .Erloes = erloes

            End With
        Catch ex As Exception
            Throw New Exception("in erstelle InventurProjekte: " & ex.Message)
        End Try



        '
        ' Ende Objekt Anlage
        '


    End Sub


    ''' <summary>
    ''' fügt dem Projekt hproj das Modul mit Namen vorlagenName hinzu
    ''' alternativ können ParentID einer Phase oder parentName und startOffset, endoffset angegeben werden
    ''' Achtung: ergänzen zu einer ParentID wird noch nicht unterstützt  - hier muss noch behandelt werden, dass
    ''' die neu hinzugefügten Phasen ja am Ende der AllPhases List angefügt werden; andererseits werden beim Darstellen des Extended Mode die Phasen eines 
    ''' nach dem anderen gezeichnet; notwendig wäre hier eine zusätzliche Struktur, die die Reihenfolge beeinhaltet oder AllPhases und die Hierarchie Struktur muss neu afgebaut bzw. geändert werden 
    ''' </summary>
    ''' <param name="hproj"></param>
    ''' <param name="parentNameID"></param>
    ''' <param name="parentName"></param>
    ''' <param name="vorlagenName"></param>
    ''' <param name="startOffset"></param>
    ''' <param name="endOffset"></param>
    ''' <remarks></remarks>
    Public Sub addModuleToProjekt(ByRef hproj As clsProjekt, ByVal vorlagenName As String, _
                                      ByVal parentNameID As String, ByVal parentName As String, _
                                      ByVal startOffset As Integer, ByVal endOffset As Integer, _
                                      ByVal dontStretch As Boolean)

        Dim modulVorlage As clsProjektvorlage
        If ModulVorlagen.Contains(vorlagenName) Then

            If parentNameID.Length > 0 Then
                Throw New ArgumentException("wird noch nicht unterstützt ...")
            ElseIf parentName.Length = 0 Then
                Throw New ArgumentException("Name muss mindestens ein Zeichen enthalten ")
            ElseIf startOffset <= 0 Or endOffset <= 0 Or endOffset - startOffset <= 0 Then
                Throw New ArgumentException("ungültige Start-/Ende Angaben: " & startOffset & " , " & endOffset)
            Else
                ' jetzt kann die Aktion durchgeführt werden 
                modulVorlage = ModulVorlagen.getProject(vorlagenName)
                modulVorlage.moduleCopyTo(project:=hproj, parentID:="", moduleName:=parentName, _
                                           modulStartOffset:=startOffset, endOffset:=endOffset, dontStretch:=dontStretch)
            End If

        Else
            Throw New ArgumentException("Vorlage " & vorlagenName & " existiert nicht")
        End If

    End Sub
        '
        '
        '
        ''' <summary>
        ''' löscht das angegebene Projekt mit Name pName inkl all seiner Varianten 
        ''' </summary>
        ''' <param name="pName">
        ''' gibt an , ob es der erste Aufruf war
        ''' wenn ja, kommt erst der Bestätigungs-Dialog 
        ''' wenn nein, wird ohne Aufforderung zur Bestätigung gelöscht 
        ''' </param>
        ''' <remarks></remarks>
    Public Sub awinDeleteProjectInSession(ByVal pName As String)


        Dim bestaetigeLoeschen As New frmconfirmDeletePrj
        Dim zeile As Integer
        Dim hproj As clsProjekt
        Dim anzahlZeilen As Integer
        Dim tmpCollection As New Collection

        Dim formerEOU As Boolean = enableOnUpdate
        enableOnUpdate = False

        hproj = ShowProjekte.getProject(pName)
        anzahlZeilen = hproj.calcNeededLines(tmpCollection, awinSettings.drawphases, False)



        If ShowProjekte.contains(pName) Then

            ' Aktuelle Konstellation ändert sich dadurch
            currentConstellation = ""

            zeile = hproj.tfZeile

            ' Shape wird gelöscht - ausserdem wird der Verweis in hproj auf das Shape gelöscht 
            Call clearProjektinPlantafel(pName)

            ShowProjekte.Remove(pName)

            ' in der Projekt-Tafel den Platz nutzen ... 
            Call moveShapesUp(zeile, anzahlZeilen)

        End If


        AlleProjekte.RemoveAllVariantsOf(pName)
        enableOnUpdate = formerEOU

    End Sub

    Public Sub awinDeleteChart(ByRef chtobj As ChartObject)
        Dim kennung As String
        Dim hDiagramm As clsDiagramm

        ' in der DiagramList wird die letzte Position gespeichert , deshlab ist es kontra produktiv , das zu löschen 

        Try
            kennung = chtobj.Name
        Catch ex As Exception
            kennung = "?"
        End Try

        Try
            hDiagramm = DiagramList.getDiagramm(kennung)
            With DiagramList.getDiagramm(kennung)
                .top = chtobj.Top
                .left = chtobj.Left
            End With
        Catch ex As Exception

        End Try

        chtobj.Delete()


    End Sub

    Public Sub awinStoreCockpit(ByVal cockpitname As String)
        'Dim kennung As String

        Dim currentDirectoryName As String = My.Computer.FileSystem.CurrentDirectory & "\"
        Dim fileName As String
        Dim found As Boolean = False
        Dim wsfound As Boolean = False
        Dim fileIsOpen As Boolean = False
        Dim anzChartsInCockpit As Integer
        Dim anzDiagrams As Integer
        Dim i As Integer
        Dim k As Integer = 1
        Dim maxRows As Integer
        Dim maxColumns As Integer
        Dim logMessage As String = " "
        Dim newchtobj As Excel.ChartObject
        Dim oldchtobj As Excel.ChartObject
        Dim chtobj As Excel.ChartObject
        Dim hchtobj As Excel.ChartObject
        Dim hshape As Excel.Shape
        Dim xlsCockpits As xlNS.Workbook = Nothing
        Dim wsSheet As xlNS.Worksheet = Nothing
        Dim wsPT As xlNS.Worksheet = Nothing
        Dim sichtbarerBereich As Excel.Range

        Try
            ' Merken des aktuell gesetzten sichtbaren Bereich in der ProjektTafel
            With appInstance.ActiveWindow
                sichtbarerBereich = .VisibleRange
            End With
            wsPT = CType(appInstance.Worksheets(arrWsNames(3)), Excel.Worksheet)
            With wsPT
                ' benötigt um die Spaltenbreite und Zeilenhöhe  zu setzen für die Tabelle in "Project Board Cockpit.xlsx", in die das neue Cockpit gespeichert wird.
                maxRows = .Rows.Count
                maxColumns = .Columns.Count
                ' Anzahl Diagramme, die gespeichert werden zu diesem Cockpit
                anzDiagrams = CType(.ChartObjects, Excel.ChartObjects).Count


                If anzDiagrams > 0 Then

                    fileName = awinPath & cockpitsFile

                    If My.Computer.FileSystem.FileExists(fileName) Then

                        Try
                            If Not fileIsOpen Then
                                xlsCockpits = appInstance.Workbooks.Open(fileName)
                                fileIsOpen = True
                            End If
                        Catch ex As Exception

                            i = 1
                            While i <= appInstance.Workbooks.Count And Not fileIsOpen
                                If appInstance.Workbooks(i).Name = fileName Then
                                    xlsCockpits = appInstance.Workbooks(i)
                                    fileIsOpen = True
                                Else
                                    i = i + 1
                                End If
                            End While

                            If Not fileIsOpen Then
                                logMessage = "Öffnen von " & fileName & " fehlgeschlagen" & vbLf & _
                                                            "falls die Datei bereits geöffnet ist: Schließen Sie sie bitte"

                                Throw New ArgumentException(logMessage)
                            End If

                        End Try
                    Else
                        ' Cockpits-File neu anlegen 
                        xlsCockpits = appInstance.Workbooks.Add()
                        xlsCockpits.SaveAs(fileName)
                    End If

                    ' wenn das richtige Tabellenblatt in Datei "Project Board Cockpits.xlsx" vorhanden, dann löschen 
                    Try
                        wsSheet = CType(xlsCockpits.Worksheets(cockpitname), Excel.Worksheet)
                        If wsSheet.Name = cockpitname Then
                            ' Tabellenblatt existiert bereits, es muss gelöscht werden und neu angelegt
                            xlsCockpits.Worksheets.Application.DisplayAlerts = False
                            wsSheet.Delete()
                            xlsCockpits.Worksheets.Application.DisplayAlerts = True
                        End If
                    Catch ex As Exception

                    End Try
                    'i = 1
                    'While i <= xlsCockpits.Worksheets.Count And Not wsfound
                    '    wsSheet = xlsCockpits.Worksheets.Item(i)
                    '    If wsSheet.Name = cockpitname Then
                    '        ' Tabellenblatt existiert bereits, es muss gelöscht werden und neu angelegt
                    '        xlsCockpits.Worksheets.Application.DisplayAlerts = False
                    '        wsSheet.Delete()
                    '        xlsCockpits.Worksheets.Application.DisplayAlerts = True
                    '        wsfound = True
                    '    Else
                    '        i = i + 1
                    '    End If

                    'End While

                    ' Tabellenblatt muss neu hinzugefügt werden

                    wsSheet = CType(xlsCockpits.Worksheets.Add(), Excel.Worksheet)
                    wsSheet.Name = cockpitname
                    ' hier werden jetzt die Spaltenbreiten und Zeilenhöhen gesetzt 

                    'With wsSheet

                    '    CType(.Range(.Cells(1, 1), .Cells(maxRows, maxColumns)), Global.Microsoft.Office.Interop.Excel.Range).RowHeight = awinSettings.zeilenhoehe2
                    '    CType(.Columns, Global.Microsoft.Office.Interop.Excel.Range).ColumnWidth = awinSettings.spaltenbreite


                    '    .Range(.Cells(2, 1), .Cells(maxRows, maxColumns)).HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter
                    '    .Range(.Cells(2, 1), .Cells(maxRows, maxColumns)).VerticalAlignment = Excel.XlVAlign.xlVAlignCenter

                    'End With

                    ' Tabellenblatt existiert jetzt sicher

                    ' alle Charts durchgehen und in "Project Board Cockpits.xlsx" Tabelle "cockpitname" speichern

                    While k <= anzDiagrams

                        wsPT.Activate()

                        chtobj = CType(wsPT.ChartObjects(k), Excel.ChartObject)

                        oldchtobj = chtobj

                        chtobj.Copy()

                        ' wenn Chart vorhanden, dann ersetzen, sonst hinzufügen

                        found = False
                        i = 1
                        anzChartsInCockpit = CType(wsSheet.ChartObjects, Excel.ChartObjects).Count
                        While i <= anzChartsInCockpit And Not found
                            hchtobj = CType(wsSheet.ChartObjects(i), Excel.ChartObject)
                            ' an awinLoadCockpit anpassen
                            If hchtobj.Name = chtobj.Name Then
                                hchtobj.Delete()
                                found = True
                            Else
                                i = i + 1
                            End If
                        End While

                        'chtobj.Cut()

                        wsSheet.Activate()

                        ' Chart aus dem Buffer nun in das Tabellenblatt einfügen
                        wsSheet.Paste()
                        anzChartsInCockpit = CType(wsSheet.ChartObjects, Excel.ChartObjects).Count

                        ' dem neu eingefügten Chart die richtige Position eintragen, neutralisiert um den sichtbaren Bereich
                        newchtobj = CType(wsSheet.ChartObjects(anzChartsInCockpit), Excel.ChartObject)

                        newchtobj.Top = oldchtobj.Top - CDbl(sichtbarerBereich.Top)
                        newchtobj.Left = oldchtobj.Left - CDbl(sichtbarerBereich.Left)


                        ' aus der DiagrammList noch DiagrammTyp herausholen und in das Chart bei AlternativText eintragen
                        Dim hdiagramm As clsDiagramm
                        i = 1
                        found = False

                        While i <= DiagramList.Count And Not found

                            hdiagramm = DiagramList.getDiagramm(i)
                            If hdiagramm.kennung = newchtobj.Name Then
                                'newchtobj.Chart.Name = hdiagramm.diagrammTyp
                                found = True
                                hshape = chtobj2shape(newchtobj)
                                hshape.Title = hdiagramm.diagrammTyp
                                Try
                                    If Not IsNothing(hdiagramm.gsCollection) Then
                                        For hi = 1 To hdiagramm.gsCollection.Count
                                            If hi = 1 Then
                                                hshape.AlternativeText = CStr(hdiagramm.gsCollection.Item(hi))
                                            Else
                                                hshape.AlternativeText = hshape.AlternativeText & ";" & CStr(hdiagramm.gsCollection.Item(hi))
                                            End If
                                        Next hi
                                    End If
                                Catch ex As Exception
                                    Throw New Exception("Fehler  Cockpits '" & cockpitname & vbLf & ex.Message)
                                End Try

                            End If
                            i = i + 1

                        End While


                        ' aus der DiagrammList noch Collection herausholen und in das Chart bei Beschreibung eintragen
                        k = k + 1

                    End While

                    appInstance.ActiveWorkbook.Close(SaveChanges:=True)

                    enableOnUpdate = True
                    appInstance.ScreenUpdating = True

                    Call MsgBox("Cockpit '" & cockpitname & "' wurde gespeichert")
                    'xlsCockpits.Close(SaveChanges:=True)

                Else
                    Call MsgBox("Es sind keine Charts vorhanden")
                End If
            End With

        Catch ex As Exception
            Throw New Exception("Fehler beim Speichern des Cockpits '" & cockpitname & vbLf & ex.Message)
        End Try

    End Sub
    Public Sub awinLoadCockpit(ByVal cockpitname As String)
        'Dim kennung As String

        Dim fileName As String
        Dim found As Boolean = False
        Dim wsfound As Boolean = False
        Dim fileIsOpen As Boolean = False
        Dim isPfDiagramm As Boolean = False
        Dim anzDiagrams As Integer
        Dim i As Integer
        Dim k As Integer = 1
        Dim j As Integer = 1
        Dim logMessage As String = " "
        Dim newchtobj As Excel.ChartObject
        Dim hchtobj As Excel.ChartObject
        Dim hshape As Excel.Shape
        Dim chtobj As Excel.ChartObject
        Dim xlsCockpits As xlNS.Workbook = Nothing
        Dim wsSheet As xlNS.Worksheet = Nothing
        Dim currentWS As xlNS.Worksheet = Nothing
        Dim sichtbarerBereich As Excel.Range
        Dim hstring As String

        appInstance.EnableEvents = False
        Try
            With appInstance.ActiveWindow
                sichtbarerBereich = .VisibleRange
            End With

            currentWS = CType(appInstance.Worksheets(arrWsNames(3)), Excel.Worksheet)

            fileName = awinPath & cockpitsFile

            If My.Computer.FileSystem.FileExists(fileName) Then

                Try

                    xlsCockpits = appInstance.Workbooks.Open(fileName)

                Catch ex As Exception

                    i = 1
                    While i <= appInstance.Workbooks.Count And Not fileIsOpen
                        If appInstance.Workbooks(i).Name = fileName Then
                            xlsCockpits = appInstance.Workbooks(i)
                            fileIsOpen = True
                        Else
                            i = i + 1
                        End If
                    End While

                    If Not fileIsOpen Then
                        logMessage = "Öffnen von " & fileName & " fehlgeschlagen" & vbLf & _
                                                    "falls die Datei bereits geöffnet ist: Schließen Sie sie bitte"

                        Throw New ArgumentException(logMessage)
                    End If

                End Try

                ' richtige Tabellenblatt  in Datei "Project Board Cockpits.xlsx" aktivieren
                Try
                    wsSheet = CType(xlsCockpits.Worksheets(cockpitname), Excel.Worksheet)

                    k = 1
                    Dim anzChartObj As Integer = CType(wsSheet.ChartObjects, Excel.ChartObjects).Count
                    While k <= anzChartObj

                        wsSheet.Activate()
                        chtobj = CType(wsSheet.ChartObjects(1), Excel.ChartObject)  ' immer das Chart 1 lesen, da die anderen mit Cut ausgeschnitten wurden

                        ' löschen der Zwischenablage
                        My.Computer.Clipboard.Clear()

                        Dim CPtmpArray() As String
                        CPtmpArray = chtobj.Name.Split(New Char() {CType("#", Char)}, 5)
                        isPfDiagramm = (CPtmpArray(0) = "pf")

                        ' testen, ob dieses Chart bereits angezeigt wird, dann ggfalls. löschen
                        found = False
                        j = 1
                        While j <= CType(currentWS.ChartObjects, Excel.ChartObjects).Count And Not found
                            hchtobj = CType(currentWS.ChartObjects(j), Excel.ChartObject)

                            Dim PTtmpArray() As String
                            PTtmpArray = hchtobj.Name.Split(New Char() {CType("#", Char)}, 5)

                            ' Überprüfen, ob das Chart bereits angezeigt wird, dann ersetzen
                            If Not IsNothing(hchtobj) Then

                                If hchtobj.Name <> "" Then

                                    If hchtobj.Name <> chtobj.Name Then
                                        ' chtobj name ist aufgebaut: pr#PTprdk.kennung#pName#Auswahl
                                        ' oder
                                        ' chtobj name ist so: pf#zahl#zahl
                                        If CPtmpArray(0) = "pr" And PTtmpArray(0) = "pr" Then
                                            If CPtmpArray(0) = PTtmpArray(0) And CPtmpArray(1) = PTtmpArray(1) And CPtmpArray(3) = PTtmpArray(3) Then
                                                currentWS.ChartObjects(j).Delete()
                                                found = True
                                            End If
                                        Else
                                            If isPfDiagramm Then
                                                If hchtobj.Name = chtobj.Name Then

                                                    currentWS.ChartObjects(j).Delete()
                                                    found = True
                                                End If
                                            End If
                                        End If
                                    Else
                                        currentWS.ChartObjects(j).Delete()
                                        found = True
                                    End If

                                End If
                            End If

                            isPfDiagramm = (CPtmpArray(0) = "pf")

                            j = j + 1
                        End While
                        ''  die Position merken
                        Dim chtTop As Double = chtobj.Top
                        Dim chtLeft As Double = chtobj.Left


                        chtobj.Cut()

                        'Dim newtestshape As Excel.ChartObject
                        currentWS.Activate()
                        currentWS.Paste()
                        anzDiagrams = CType(currentWS.ChartObjects, Excel.ChartObjects).Count
                        ' dem neu eingefügten Chart die richtige Position eintragen
                        newchtobj = CType(currentWS.ChartObjects(anzDiagrams), Excel.ChartObject)
                        newchtobj.Top = CDbl(sichtbarerBereich.Top) + chtTop
                        newchtobj.Left = CDbl(sichtbarerBereich.Left) + chtLeft

                        If isPfDiagramm Then

                            ' Alternativtext herausbekommen
                            hshape = chtobj2shape(newchtobj)


                            ' hier wird die maximale Anzahl an Phasen oder Rollen oder Kosten herausgefunden
                            Dim maxAnz As Integer = System.Math.Max(RoleDefinitions.Count, PhaseDefinitions.Count)
                            maxAnz = System.Math.Max(maxAnz, CostDefinitions.Count)

                            Dim tmpArray1() As String
                            Dim myCollection As New Collection
                            tmpArray1 = hshape.AlternativeText.Split(New Char() {CType(";", Char)}, maxAnz)

                            ' myCollection aufbauen mit den verschiedenen Werten die im Diagramm angezeigt werden sollen
                            For hi = 0 To tmpArray1.Length - 1
                                If tmpArray1.Length >= 1 And tmpArray1(hi) <> "" Then
                                    hstring = tmpArray1(hi)
                                    myCollection.Add(hstring, hstring)
                                    'Else
                                    '    myCollection = Nothing
                                End If

                            Next hi

                            ' Diagramme in die diagrammListe einfügen mit allen Angaben

                            Dim prcDiagram As New clsDiagramm

                            ' Anfang Event Handling für Chart 
                            Dim prcChart As New clsEventsPrcCharts
                            prcChart.PrcChartEvents = newchtobj.Chart
                            prcDiagram.setDiagramEvent = prcChart
                            ' Ende Event Handling für Chart 

                            With prcDiagram
                                .DiagrammTitel = newchtobj.Chart.ChartTitle.Text
                                .diagrammTyp = hshape.Title.Trim
                                .gsCollection = myCollection
                                .isCockpitChart = False
                                .top = newchtobj.Top
                                .left = newchtobj.Left
                                .width = newchtobj.Width
                                .height = newchtobj.Height
                                .kennung = newchtobj.Name
                            End With

                            ' eintragen in die sortierte Liste mit .kennung als dem Schlüssel 
                            ' wenn das Diagramm bereits existiert, muss es gelöscht werden, dann neu ergänzt ... 
                            Try
                                DiagramList.Add(prcDiagram)
                            Catch ex As Exception
                                Try
                                    DiagramList.Remove(prcDiagram.kennung)
                                    DiagramList.Add(prcDiagram)
                                Catch ex1 As Exception

                                End Try
                            End Try

                        End If                ' Ende der PF-Diagramm Spezialbehandlung

                        k = k + 1

                    End While
                    wsfound = True
                    'Call MsgBox("Es wurden " & k - 1 & " Charts eingefügt")
                Catch ex As Exception
                    xlsCockpits.Close(SaveChanges:=False)
                    Throw New ArgumentException("Fehler beim Laden des Cockpits '" & cockpitname & vbLf, ex.Message)
                End Try

                xlsCockpits.Close(SaveChanges:=False)

            Else
                ' Project Board Cockpit.xlsx ist nicht vorhanden
                Call MsgBox("Es sind keine Charts vorhanden." & vbLf & "'Project Board Cockpit.xlsx ist nicht vorhanden.")

            End If

        Catch ex As Exception
            xlsCockpits.Close(SaveChanges:=False)
            Throw New ArgumentException("Fehler beim Laden des Cockpits '" & cockpitname & vbLf, ex.Message)
        End Try
        appInstance.EnableEvents = True
    End Sub
    '
    ''' <summary>
    ''' gibt die Referenz für dazu zum ChartObject cho gehörige Excel.Shape zurück
    ''' </summary>
    ''' <param name="cho">ChartObject</param>
    ''' <remarks></remarks>
    Function chtobj2shape(ByRef cho As ChartObject) As Excel.Shape

        'Dim zo As Long
        Dim ws As Excel.Worksheet
        Dim shc As Excel.Shapes
        Dim sh As Excel.Shape
        Dim found As Boolean = False
        Dim i As Integer = 0
        ws = CType(cho.Parent, Excel.Worksheet)
        shc = ws.Shapes
        While Not found And i <= shc.Count
            i = i + 1
            found = cho.Name = shc.Item(i).Name
        End While

        sh = shc.Item(i)

        chtobj2shape = sh

    End Function
    '
    ''' <summary>
    ''' das Projekt beauftragen bzw. die Änderungen akzeptieren
    ''' type 0 : Acceptchanges
    ''' type 1: Beauftragen 
    ''' </summary>
    ''' <param name="pname">Projektname</param>
    ''' <param name="type">0: Accept Changes; 1: Beauftragung </param>
    ''' <remarks></remarks>
    Public Sub awinBeauftragung(ByVal pname As String, ByVal type As Integer)
        Dim hproj As clsProjekt
        Dim zeile As Integer


        ' prüfen, ob es in der ShowProjektListe ist ...
        If ShowProjekte.contains(pname) Then


            Try
                hproj = ShowProjekte.getProject(pname)
                With hproj
                    zeile = .tfZeile
                    If type = 1 Then
                        .Status = ProjektStatus(1)
                    End If
                    .diffToPrev = False
                    .timeStamp = Date.Now
                End With

                ' wenn bestimmte Projekte beim Suchen nach einem Platz nicht berücksichtigt werden sollen,
                ' dann müssen sie in einer Collection an ZeichneProjektinPlanTafel übergeben werden 
                Dim tmpCollection As New Collection
                Call ZeichneProjektinPlanTafel(tmpCollection, pname, zeile, tmpCollection, tmpCollection)

            Catch ex As Exception
                Call MsgBox(" Fehler in Beauftragung " & pname & " , Modul: awinBeauftragung")
                Exit Sub
            End Try


        Else
            Call MsgBox("Projekt " & pname & " wurde nicht gefunden")
        End If

    End Sub


    Public Sub awinCancelBeauftragung(ByVal pname As String)
        Dim hproj As clsProjekt
        Dim zeile As Integer


        ' prüfen, ob es in der ShowProjektListe ist ...
        If ShowProjekte.contains(pname) Then


            Try
                hproj = ShowProjekte.getProject(pname)

                If hproj.variantName = "" Then
                    Call MsgBox("die Fixierung der Standard Variante kann nicht aufgehoben werden ..." & vbLf & _
                                "bitte erstellen Sie zu diesem Zweck eine Variante ...")
                Else
                    With hproj
                        zeile = .tfZeile
                        .Status = ProjektStatus(0)
                        .timeStamp = Date.Now
                    End With


                    ' wenn bestimmte Projekte beim Suchen nach einem Platz nicht berücksichtigt werden sollen,
                    ' dann müssen sie in einer Collection an ZeichneProjektinPlanTafel übergeben werden 
                    Dim tmpCollection As New Collection
                    Call ZeichneProjektinPlanTafel(tmpCollection, pname, zeile, tmpCollection, tmpCollection)

                End If

            Catch ex As Exception
                Call MsgBox(" Fehler in Fixierung aufheben " & pname & " , Modul: awinCancelBeauftragung")
                Exit Sub
            End Try


        Else
            Call MsgBox("Projekt " & pname & " wurde nicht gefunden")
        End If

    End Sub
    ''' <summary>
    ''' stellt das Projekt "pname" ims NoShow
    ''' </summary>
    ''' <param name="pname">NAme des Projekts</param>
    ''' <remarks></remarks>
    Public Sub awinShowNoShowProject(ByVal pname As String)
        Dim hproj As clsProjekt

        'Dim tfz As Integer, tfs As Integer



        ' prüfen, ob es in der ShowProjektListe ist ...
        If ShowProjekte.contains(pname) Then

            ' Shape wird gelöscht - ausserdem wird der Verweis in hproj auf das Shape gelöscht 
            Call clearProjektinPlantafel(pname)


            Try
                hproj = ShowProjekte.getProject(pname)
                ' aus Showprojekte rausnehmen
                ShowProjekte.Remove(pname)

                ' ist es bereits eine andere Variante in NoShowPRojekte?
                If noShowProjekte.contains(pname) Then
                    noShowProjekte.Remove(pname)
                End If
                noShowProjekte.Add(hproj)

            Catch ex As Exception
                Call MsgBox(" Fehler in NoShow " & pname & " , Modul: NoShowProject")
                Exit Sub
            End Try




            'Dim abstand As Integer ' eigentlich nur Dummy Variable, wird aber in Tabelle2 benötigt ...
            'Call awinClkReset(abstand)

            ' ein Projekt wurde gelöscht bzw aus Showprojekte entfernt  - typus = 3
            Call awinNeuZeichnenDiagramme(3)



        Else
            Call MsgBox("Projekt " & pname & " wurde nicht gefunden")
        End If

    End Sub
    ''' <summary>
    ''' vergleicht das selektierte Projekt, in Abhängigkeit von compareType mit 
    ''' dem Projekt-Template, dem Beauftragungs-Stand, einer Projekt-Variante oder einem anderen Projekt
    ''' compareType = 0 Projekt-Template
    ''' compareType = 1 Projekt-Variante
    ''' compareType = 2 Beauftragung 
    ''' compareType = 3 anderes Projekt
    ''' compareType = 4 zwei zeitliche Versionen eines Projektes 
    ''' </summary>
    ''' <param name="compareType"></param>
    ''' <remarks></remarks>
    Public Sub awinCompareProject(ByRef hproj As clsProjekt, ByRef cproj As clsProjekt, ByVal compareType As Integer, ByVal top As Double, ByVal left As Double)

        'Dim vname As String

        Dim Werte1() As Double, Werte2() As Double
        Dim listeRollen As New Collection
        Dim listeKosten As New Collection
        Dim listeTemp As New Collection
        Dim hvproj As New clsProjektvorlage
        Dim i As Integer
        Dim width As Double, height As Double
        Dim atLeastOne As Boolean = False
        Dim awinMessage As String = " ... es sind keine Unterschiede festzustellen  ..."
        Dim pname2 As String = cproj.name
        Dim pname1 As String = hproj.name
        Dim titel1 As String = pname1, titel2 As String = pname2


        If (pname1 = pname2) And hproj.variantName = cproj.variantName Then
            If compareType = 3 Then
                ' Vergleich mit Beauftragung
                titel1 = pname1 & " (" & hproj.timeStamp.ToString & ")"
                titel2 = "Beauftragung (" & cproj.timeStamp.ToString & ")"
            ElseIf compareType = 4 Then
                ' Vergleich mit Beauftragung
                titel1 = pname1 & " (" & hproj.timeStamp.ToString & ")"
                titel2 = "akt. Stand (" & cproj.timeStamp.ToString & ")"
            End If
            ' die verschiedenen Planungs-Stände eines Projektes werden miteinander verglichen 
        End If


        ' als erstes werden alle entweder in aktueller Stand bzw in Vorlage vorhandenen Rollen/Kostenarten gesucht  
        ' und werden in listeRollen bzw listeKosten gespeichert

        listeRollen = hproj.getUsedRollen
        listeTemp = cproj.getUsedRollen
        For i = 1 To listeTemp.Count
            Try
                If Not listeRollen.Contains(CStr(listeTemp.Item(i))) Then
                    listeRollen.Add(listeTemp.Item(i))
                End If
            Catch ex As Exception

            End Try

        Next

        listeKosten = hproj.getUsedKosten
        listeTemp = cproj.getUsedKosten

        For i = 1 To listeTemp.Count
            Try
                If Not listeKosten.Contains(CStr(listeTemp.Item(i))) Then
                    listeKosten.Add(listeTemp.Item(i))
                End If
            Catch ex As Exception

            End Try

        Next


        ' neu 
        '
        ' die Position des Diagramms wird ausgerechnet ...
        '



        height = 180
        width = hproj.anzahlRasterElemente * boxWidth + 10


        Dim hname As String, mEinheit As String = awinSettings.kapaEinheit


        ' jetzt werden für alle  Rollenbedarfe, sofern unterschiedlich die Diagramme gezeichnet ... 
        Try
            For i = 1 To listeRollen.Count
                hname = CStr(listeRollen.Item(i))
                Werte1 = hproj.getRessourcenBedarf(hname)
                Werte2 = cproj.getRessourcenBedarf(hname)
                If arraysAreDifferent(Werte1, Werte2) Then
                    ' if absolut then 
                    Call ShowDiagramCompare(titel1, Werte1, titel2, Werte2, hname, mEinheit, top, left, width, height)
                    'Else
                    '    Call ShowDiagramCompare2(name1, Werte1, name2, Werte2, hname, mEinheit, top, left, width, height)
                    'End If
                    atLeastOne = True
                    top = top + 10
                    left = left + 10
                End If

            Next
        Catch ex As Exception

        End Try


        ' jetzt werden für alle  Kostenbedarfe, sofern unterschiedlich die Diagramme gezeichnet ... 

        mEinheit = "T€"
        Try
            For i = 1 To listeKosten.Count
                hname = CStr(listeKosten.Item(i))
                Werte1 = hproj.getKostenBedarf(hname)
                Werte2 = cproj.getKostenBedarf(hname)
                If arraysAreDifferent(Werte1, Werte2) Then

                    Call ShowDiagramCompare(titel1, Werte1, titel2, Werte2, hname, mEinheit, top, left, width, height)
                    'Else
                    '    Call ShowDiagramCompare2(name1, Werte1, name2, Werte2, hname, mEinheit, top, left, width, height)
                    'End If
                    atLeastOne = True
                    top = top + 10
                    left = left + 10
                End If

            Next
        Catch ex As Exception

        End Try


        ' jetzt wird für die Gesamt-Kosten, sofern unterschiedlich das Diagramm gezeichnet ... 

        mEinheit = "T€"
        hname = "Gesamtkosten"
        Werte1 = hproj.getGesamtKostenBedarf
        Werte2 = cproj.getGesamtKostenBedarf
        If arraysAreDifferent(Werte1, Werte2) Then
            'If absolut Then
            Try
                Call ShowDiagramCompare(titel1, Werte1, titel2, Werte2, hname, mEinheit, top, left, width, height)
                'Else
                '    Call ShowDiagramCompare2(name1, Werte1, name2, Werte2, hname, mEinheit, top, left, width, height)
                'End If
                atLeastOne = True
                top = top + 10
                left = left + 10
            Catch ex As Exception
                awinMessage = ex.Message
                atLeastOne = False
            End Try

        End If

        If Not atLeastOne Then
            Call MsgBox(awinMessage)
        End If

    End Sub

    ''' <summary>
    ''' vergleicht die Phasen der beiden übergebenen Projekte, in Abhängigkeit von compareType mit 
    ''' dem Projekt-Template, dem Beauftragungs-Stand, einer Projekt-Variante oder einem anderen Projekt
    ''' compareType = 0 Projekt-Template, name2= ""
    ''' compareType = 1 Projekt-Variante, name2=varianten-Name
    ''' compareType = 2 Beauftragung , name=""
    ''' compareType = 3 anderes Projekt, name=Name des Projektes
    ''' </summary>
    ''' <param name="hproj">home-Projekt</param>
    ''' <param name="htitel">Bezeichnung Home Projekt</param>
    ''' <param name="cproj">compare Projekt</param>
    ''' <param name="ctitel">Bezeichnung compare Projekt</param>
    ''' <param name="compareType">
    ''' 0 - mit template vergleichen
    ''' 1 - Projekt-Variante vergleichen
    ''' 2 - mit anderer zeitlicher Version des gleichen Projekts 
    ''' 3 - anderes Projekt
    ''' </param>
    ''' <remarks></remarks>
    Public Sub awinCompareProjectPhases(ByVal hproj As clsProjekt, ByVal htitel As String, _
                                        ByVal cproj As clsProjekt, ByVal ctitel As String, _
                                        ByVal compareType As Integer, _
                                        ByRef chtobj As Excel.ChartObject)

        Dim vname As String, phaseNameID As String
        'Dim Werte1() As Double, Werte2() As Double
        Dim listeRollen As New Collection
        Dim listeKosten As New Collection
        Dim listeTemp As New Collection
        Dim hvproj As New clsProjektvorlage
        Dim i As Integer
        Dim top As Double, left As Double, width As Double, height As Double
        Dim atLeastOne As Boolean = False

        Dim mxAnzPhasen As Integer
        Dim tdatenreihe1() As Double, tdatenreihe2() As Double
        Dim tdatenreihe3() As Double, tdatenreihe4() As Double
        Dim tdatenreihe5() As Double, tdatenreihe6() As Double, tdatenreihe7() As Double

        Dim diagramTitle As String
        Dim anzDiagrams As Integer
        Dim found As Boolean
        Dim plen As Integer

        Dim Xdatenreihe() As String
        Dim valueColor() As Object

        Dim chtTitle As String
        'Dim pstart As Integer
        Dim ergListe As New Collection

        ' der Text, der in der Legende steht
        Dim hlegendText As String, clegendText As String

        ' je nach compare Typ werden jetzt die ChartTitel und Legendentexte gesetzt ...
        Try

            Select Case compareType
                Case 0 ' Vergleich mit Vorlage
                    vname = hproj.VorlagenName

                    hvproj = Projektvorlagen.getProject(vname)
                    hvproj.copyTo(cproj)
                    cproj.name = "Vorlage " & vname.Trim
                    ctitel = cproj.name

                    diagramTitle = "Vergleich mit Vorlage"
                    hlegendText = hproj.name & "( " & hproj.timeStamp.ToShortDateString & " )"
                    clegendText = cproj.name

                Case 1 ' Vergleich mit Projekt-Variante
                    diagramTitle = "Vergleich von zwei Projekt-Varianten"
                    hlegendText = htitel & ", Variante: " & hproj.variantName
                    clegendText = ctitel & ", Variante: " & cproj.variantName

                Case 2 ' Vergleich mit anderer zeitlicher Version 
                    diagramTitle = "Vergleich zeitlicher Versionen" & vbLf & hproj.name
                    hlegendText = htitel
                    clegendText = ctitel

                Case 3 ' Vergleich mit anderem Projekt

                    diagramTitle = "Vergleich von zwei Projekten"
                    hlegendText = htitel
                    clegendText = ctitel

                Case Else

                    Throw New ArgumentException("keine Aktion definiert für Comparetype = " & compareType)

            End Select



        Catch ex As Exception

            Throw New Exception("Fehler in compareProjectPhases: " & ex.Message)

        End Try







        ' in ErgListe werden jetzt alle Phasen-Namen aufgeführt, die entweder in cproj oder in hproj vorkommen 

        For Each cphase In hproj.Liste

            If Not ergListe.Contains(cphase.nameID) Then
                ergListe.Add(cphase.nameID, cphase.nameID)
            End If

        Next









        ' in cproj könnten ja Phasen auftauchen, die in hproj nicht drin sind ...
        For Each cphase In cproj.Liste

            If Not ergListe.Contains(cphase.nameID) Then
                ergListe.Add(cphase.nameID, cphase.nameID)
            End If

        Next


        mxAnzPhasen = ergListe.Count


        appInstance.EnableEvents = False
        appInstance.ScreenUpdating = False

        '
        ' hole die größere der beiden Projektdauern
        '

        plen = System.Math.Max(hproj.dauerInDays, cproj.dauerInDays)


        ReDim Xdatenreihe(mxAnzPhasen - 1)
        ReDim tdatenreihe1(mxAnzPhasen - 1)
        ReDim tdatenreihe2(mxAnzPhasen - 1)
        ReDim tdatenreihe3(mxAnzPhasen - 1)
        ReDim tdatenreihe4(mxAnzPhasen - 1)
        ReDim tdatenreihe5(mxAnzPhasen - 1)
        ReDim tdatenreihe6(mxAnzPhasen - 1)
        ReDim tdatenreihe7(mxAnzPhasen - 1)


        ReDim valueColor(mxAnzPhasen - 1)

        For i = 1 To mxAnzPhasen
            phaseNameID = CStr(ergListe.Item(i))
            Xdatenreihe(i - 1) = elemNameOfElemID(phaseNameID)
        Next i



        'ReDim hsum(anzPhasen - 1)

        Dim p1(1) As Integer, p2(1) As Integer
        Dim vgl(1) As Integer, vgl2(1) As Integer
        Dim p1Vorp2 As Boolean

        For i = 1 To mxAnzPhasen
            ReDim p1(1)
            ReDim p2(1)
            ReDim vgl(1)
            ReDim vgl2(1)
            p1Vorp2 = False
            phaseNameID = CStr(ergListe.Item(i))

            Try
                With hproj.getPhaseByID(phaseNameID)
                    p1(0) = .startOffsetinDays
                    p1(1) = .startOffsetinDays + .dauerInDays - 1
                End With

                Try
                    With cproj.getPhaseByID(phaseNameID)
                        p2(0) = .startOffsetinDays
                        p2(1) = .startOffsetinDays + .dauerInDays - 1
                    End With
                    ' sowohl Projekt 1 als auch Projekt 2 haben die Phase 
                    '
                    ' Bestimmen tdatenreihe1 - Null-Farbe von Anfang bis Start P1 bzw Start P2 
                    vgl(0) = p1(0)
                    vgl(1) = p2(0)
                    tdatenreihe1(i - 1) = vgl.Min - 1

                    ' bestimmen tdatenreihe2(i-1) - wenn p1 vor p2 steht 
                    tdatenreihe2(i - 1) = p2(0) - p1(0)
                    If tdatenreihe2(i - 1) < 0 Then
                        p1Vorp2 = False
                        tdatenreihe2(i - 1) = 0
                    End If

                    ' bestimmen tdatenreihe3(i-1) - wenn p2 vor p1 steht 
                    tdatenreihe3(i - 1) = p1(0) - p2(0)
                    If tdatenreihe3(i - 1) < 0 Then
                        p1Vorp2 = True
                        tdatenreihe3(i - 1) = 0
                    End If

                    ' bestimmen tdatenreihe4 - wie groß ist die Lücke zwischen beiden, wenn sie sich nicht überlappen .. 
                    If p1Vorp2 Then
                        tdatenreihe4(i - 1) = p2(0) - p1(1) - 1
                    Else
                        tdatenreihe4(i - 1) = p1(0) - p2(1) - 1
                    End If

                    If tdatenreihe4(i - 1) < 0 Then
                        tdatenreihe4(i - 1) = 0
                    End If


                    ' bestimmen tdatenreihe5 - wieviel haben beide gemeinsam 
                    vgl(0) = p1(0)
                    vgl(1) = p2(0)
                    vgl2(0) = p1(1)
                    vgl2(1) = p2(1)

                    tdatenreihe5(i - 1) = vgl2.Min - vgl.Max + 1
                    If tdatenreihe5(i - 1) < 0 Then
                        tdatenreihe5(i - 1) = 0
                    End If


                    ' bestimmen tdatenreihe6 - der Teil, um den P1 länger dauert als P2  
                    tdatenreihe6(i - 1) = p1(1) - p2(1)
                    If tdatenreihe6(i - 1) < 0 Then
                        tdatenreihe6(i - 1) = 0
                    End If

                    ' bestimmen tdatenreihe7 - der Teil, um den P2 länger dauert als P1  
                    tdatenreihe7(i - 1) = p2(1) - p1(1)
                    If tdatenreihe7(i - 1) < 0 Then
                        tdatenreihe7(i - 1) = 0
                    End If

                Catch ex1 As Exception

                    ' Projekt2 hat die Phase nicht ...
                    tdatenreihe1(i - 1) = p1(0) - 1
                    tdatenreihe5(i - 1) = p1(1) - p1(0) + 1

                End Try
            Catch ex As Exception

                ' Projekt1 hat die Phase nicht ...
                Try
                    With cproj.getPhaseByID(phaseNameID)
                        p2(0) = .startOffsetinDays
                        p2(1) = .startOffsetinDays + .dauerInDays - 1
                    End With

                    ' Projekt2 hat die Phase, Projekt1 hat sie nicht 
                    tdatenreihe1(i - 1) = p2(0) - 1
                    tdatenreihe6(i - 1) = p2(1) - p2(0) + 1


                Catch ex2 As Exception
                    ' weder Projekt1 noch PRojekt 2 hat die Phase - kann eigentlich nicht sein
                    ' in der ergListe sind per-Definitionem nur Namen, die mind in einem vorkoammen 
                    Throw New ArgumentException("in awinCompareProjectPhases: " & ex2.Message)
                    Exit Sub
                End Try


            End Try

        Next i

        '
        ' die Position des Diagramms wird ausgerechnet ...
        '

        top = 48 + hproj.tfZeile * 15
        left = (hproj.tfspalte - 1) * boxWidth - 5

        If left < 0 Then
            left = 5
        End If

        height = (mxAnzPhasen - 1) * 20 + 90
        width = plen / 365 * 12 * boxWidth + 10




        With CType(appInstance.Worksheets(arrWsNames(3)), Excel.Worksheet)
            anzDiagrams = CType(.ChartObjects, Excel.ChartObjects).Count
            '
            ' um welches Diagramm handelt es sich ...
            '
            i = 1
            found = False
            While i <= anzDiagrams And Not found
                Try
                    chtTitle = CType(.ChartObjects(i), Excel.ChartObject).Chart.ChartTitle.Text
                Catch ex As Exception
                    chtTitle = " "
                End Try
                If chtTitle = diagramTitle Then
                    found = True
                Else
                    i = i + 1
                End If

            End While

            If found Then
                MsgBox(" Diagramm wird bereits angezeigt ...")
            Else

                With appInstance.Charts.Add
                    ' remove extra series
                    Do Until .SeriesCollection.Count = 0
                        .SeriesCollection(1).Delete()
                    Loop

                    'Aufbau der Series 
                    With .SeriesCollection.NewSeries
                        .name = "null"
                        .Interior.colorindex = -4142
                        .Values = tdatenreihe1
                        .XValues = Xdatenreihe
                        .ChartType = Excel.XlChartType.xlBarStacked
                    End With

                    With .SeriesCollection.NewSeries
                        .name = "P1 steht vor P2"
                        .Interior.color = vergleichsfarbe1
                        .Values = tdatenreihe2
                        .XValues = Xdatenreihe
                        .ChartType = Excel.XlChartType.xlBarStacked
                    End With

                    With .SeriesCollection.NewSeries
                        .name = "P2 steht vor P1"
                        .Interior.color = vergleichsfarbe2
                        .Values = tdatenreihe3
                        .XValues = Xdatenreihe
                        .ChartType = Excel.XlChartType.xlBarStacked
                    End With

                    With .SeriesCollection.NewSeries
                        .name = "Lücke zwischen P1 und P2"
                        .Interior.colorindex = -4142
                        .Values = tdatenreihe4
                        .XValues = Xdatenreihe
                        .ChartType = Excel.XlChartType.xlBarStacked
                    End With

                    With .SeriesCollection.NewSeries
                        .name = "identisch"
                        .Interior.color = vergleichsfarbe0
                        .Values = tdatenreihe5
                        .XValues = Xdatenreihe
                        .ChartType = Excel.XlChartType.xlBarStacked
                    End With

                    With .SeriesCollection.NewSeries
                        .name = "P1 endet nach P2"
                        .Interior.color = vergleichsfarbe1
                        .Values = tdatenreihe6
                        .XValues = Xdatenreihe
                        .ChartType = Excel.XlChartType.xlBarStacked
                    End With

                    With .SeriesCollection.NewSeries
                        .name = "P2 endet nach P1"
                        .Interior.color = vergleichsfarbe2
                        .Values = tdatenreihe7
                        .XValues = Xdatenreihe
                        .ChartType = Excel.XlChartType.xlBarStacked
                    End With


                    .HasAxis(Excel.XlAxisType.xlCategory) = True
                    .HasAxis(Excel.XlAxisType.xlValue) = True

                    With .Axes(Excel.XlAxisType.xlCategory)
                        .HasTitle = False
                        .ReversePlotOrder = True
                        '.MinimumScale = 0
                        'With .AxisTitle
                        '    .Characters.text = "Phasen"
                        '    .Font.Size = 8
                        'End With
                    End With

                    With .Axes(Excel.XlAxisType.xlValue)
                        .HasTitle = True
                        .MinimumScale = 0
                        .MaximumScale = plen

                        With .AxisTitle
                            .Characters.text = "Tage"
                            .Font.Size = 8
                        End With
                    End With

                    .HasLegend = False

                    'With .Legend
                    '    .Position = XlConstants.xlTop
                    '    .Font.Size = 8
                    'End With

                    .HasTitle = True
                    .ChartTitle.Text = diagramTitle
                    .ChartTitle.font.size = 10
                    .Location(Where:=XlChartLocation.xlLocationAsObject, Name:=appInstance.Worksheets(arrWsNames(3)).name)
                End With

                ' jetzt kommt die Korrektur der Größe; herausfinden, wieviel Raum die Axis Beschriftung einnimmt ... 
                With .ChartObjects(anzDiagrams + 1)
                    .top = top
                    .height = mxAnzPhasen * 20 + 70

                    Dim axCleft As Double, axCwidth As Double
                    If .Chart.HasAxis(Excel.XlAxisType.xlCategory) = True Then
                        With CType(.Chart.Axes(Excel.XlAxisType.xlCategory), Excel.Axis)
                            axCleft = .Left
                            axCwidth = .Width
                        End With
                        If left - axCwidth < 1 Then
                            .left = 1
                            .width = width + left + 9
                        Else
                            .left = left - axCwidth
                            .width = width + axCwidth + 9
                        End If
                    Else
                        .left = left
                        .width = width
                    End If



                End With


            End If


        End With


        'Call awinScrollintoView()
        appInstance.EnableEvents = True
        appInstance.ScreenUpdating = True



    End Sub

    ''' <summary>
    ''' gibt zurück, ob die beiden sortierten Listen vom Datum her identisch sind; 
    ''' die Namen werden nicht berücksichtigt 
    ''' </summary>
    ''' <param name="liste1"></param>
    ''' <param name="liste2"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function dateListsareDifferent(ByRef liste1 As SortedList(Of Date, String), _
                                              ByRef liste2 As SortedList(Of Date, String)) As Boolean
        Dim isDifferent As Boolean = False
        Dim anzItems As Integer = liste1.Count

        If liste1.Count <> liste2.Count Then
            isDifferent = True
        Else
            Dim i As Integer = 1
            Do While i <= anzItems And Not isDifferent
                If DateDiff(DateInterval.Day, liste1.ElementAt(i - 1).Key, liste2.ElementAt(i - 1).Key) <> 0 Then
                    isDifferent = True
                End If
                i = i + 1
            Loop
        End If


        dateListsareDifferent = isDifferent

    End Function
    ''' <summary>
    ''' prüft ob zwei Arrays sowohl in der Länge als auch in den Werten absolut identisch sind
    ''' </summary>
    ''' <param name="values1"></param>
    ''' <param name="values2"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function arraysAreDifferent(ByRef values1() As Double, ByRef values2() As Double) As Boolean

        Dim istIdentisch As Boolean = True
        Dim i As Integer


        Try

            If values1.Length <> values2.Length Then
                istIdentisch = False
            End If
            i = 0
            While i <= values1.Length - 1 And istIdentisch
                If values1(i) <> values2(i) Then
                    istIdentisch = False
                Else
                    i = i + 1
                End If
            End While

        Catch ex As Exception
            Call MsgBox(ex.Message & " in arraysAreDifferent")
        End Try

        arraysAreDifferent = Not istIdentisch


    End Function


    ''' <summary>
    ''' Prozedur zeigt die Ressourcen Bedarfe des Projektes an
    ''' Auswahl = 1 : zeige den Ressourcen Bedarf
    ''' Auswahl = 2 : zeige den Personalkosten-Bedarf
    ''' Auswahl = 3 : zeige den Gesamt-Kostenbedarf
    ''' </summary>
    ''' <param name="pname">der Projekt-Name</param>
    ''' <param name="auswahl">steuert, was gezeigt wird: Ressourcen (1), Personal-Kosten (2) oder Gesamtkosten (3) </param>
    ''' <remarks></remarks>
    Public Sub ShowDiagramsRess(ByRef projektHistorie As SortedList(Of String, clsProjekt), ByVal pname As String, ByVal auswahl As Integer)
        Dim diagramTitle As String
        Dim anzDiagrams As Integer
        Dim anzRollen As Integer
        Dim found As Boolean
        Dim r As Integer, plen As Integer
        Dim i As Integer
        Dim Xdatenreihe() As String
        Dim tdatenreihe() As Double
        Dim hsum() As Double, gesamt_summe As Double
        Dim top As Double, left As Double, width As Double, height As Double
        Dim chtTitle As String
        Dim pstart As Integer
        Dim hproj As clsProjekt
        Dim ErgebnisListeR As New Collection
        Dim roleName As String
        Dim zE As String = "(" & awinSettings.kapaEinheit & ")"

        appInstance.EnableEvents = False
        appInstance.ScreenUpdating = False

        If auswahl = 1 Then
            diagramTitle = "Ressourcen-Bedarf " & zE & vbLf & pname
        ElseIf auswahl = 2 Then
            diagramTitle = "Personalkosten (T€)" & vbLf & pname
        Else
            diagramTitle = "Gesamt " & pname
        End If

        hproj = ShowProjekte.getProject(pname)
        'projektHistorie = retrie

        '
        ' hole die Projektdauer
        '
        With hproj
            plen = .anzahlRasterElemente
            pstart = .Start
        End With

        '
        ' hole die Anzahl Rollen, die in diesem Projekt vorkommen
        '
        ErgebnisListeR = hproj.getUsedRollen
        anzRollen = ErgebnisListeR.Count

        If anzRollen = 0 Then
            MsgBox("keine Ressourcen-Bedarfe definiert")
            Exit Sub
        End If


        ReDim Xdatenreihe(plen - 1)
        ReDim tdatenreihe(plen - 1)


        ReDim hsum(anzRollen - 1)

        'For m = 0 To plen - 1
        '    Xdatenreihe(m) = "" & m + 1
        'Next m

        For i = 1 To plen
            Xdatenreihe(i - 1) = StartofCalendar.AddMonths(pstart + i - 2).ToString("MMM yy")
        Next i


        ' neu 
        '
        ' die Position des Diagramms wird ausgerechnet ...
        '
        top = 48 + (hproj.tfZeile) * 15
        left = (hproj.tfspalte - 1) * boxWidth - 5
        If left < 0 Then
            left = 0
        End If
        height = awinSettings.ChartHoehe1
        width = plen * boxWidth + 10

        gesamt_summe = 0

        With CType(appInstance.Worksheets(arrWsNames(3)), Excel.Worksheet)
            anzDiagrams = CType(.ChartObjects, Excel.ChartObjects).Count
            '
            ' um welches Diagramm handelt es sich ...
            '
            i = 1
            found = False
            While i <= anzDiagrams And Not found
                Try
                    chtTitle = CType(.ChartObjects(i), Excel.ChartObject).Chart.ChartTitle.Text
                Catch ex As Exception
                    chtTitle = " "
                End Try
                If chtTitle = diagramTitle Then
                    found = True
                Else
                    i = i + 1
                End If

            End While

            If found Then
                MsgBox(" Diagramm wird bereits angezeigt ...")
            Else

                With appInstance.Charts.Add
                    ' remove extra series
                    Do Until .SeriesCollection.Count = 0
                        .SeriesCollection(1).Delete()
                    Loop



                    For r = 1 To anzRollen
                        roleName = CStr(ErgebnisListeR.Item(r))
                        If auswahl = 1 Then
                            tdatenreihe = hproj.getRessourcenBedarf(roleName)
                        Else
                            tdatenreihe = hproj.getPersonalKosten(roleName)
                        End If
                        hsum(r - 1) = 0
                        For i = 0 To plen - 1
                            hsum(r - 1) = hsum(r - 1) + tdatenreihe(i)
                        Next i
                        gesamt_summe = gesamt_summe + hsum(r - 1)

                        'series
                        With .SeriesCollection.NewSeries
                            .name = roleName
                            .Interior.color = RoleDefinitions.getRoledef(roleName).farbe
                            .Values = tdatenreihe
                            .XValues = Xdatenreihe
                            .ChartType = Excel.XlChartType.xlColumnStacked
                        End With

                    Next r

                    .HasAxis(Excel.XlAxisType.xlCategory) = True
                    .HasAxis(Excel.XlAxisType.xlValue) = True

                    With .Axes(Excel.XlAxisType.xlCategory)
                        .HasTitle = False
                        '.MinimumScale = 0
                        'With .AxisTitle
                        '    .Characters.text = "Monate"
                        '    .Font.Size = 8
                        'End With
                    End With

                    With .Axes(Excel.XlAxisType.xlValue)
                        .HasTitle = False
                        .MinimumScale = 0

                        'With .AxisTitle
                        '    If auswahl = 1 Then
                        '        .Characters.text = "Ressourcen"
                        '    Else
                        '        .Characters.text = "Personalkosten"
                        '    End If
                        '    .Font.Size = 8
                        'End With
                    End With

                    .HasLegend = True
                    With .Legend
                        .Position = Excel.Constants.xlTop
                        .Font.Size = awinSettings.fontsizeItems
                    End With
                    .HasTitle = True
                    .ChartTitle.Text = diagramTitle
                    .ChartTitle.font.size = awinSettings.fontsizeTitle
                    .Location(Where:=XlChartLocation.xlLocationAsObject, Name:=appInstance.Worksheets(arrWsNames(3)).name)
                End With

                ' jetzt kommt die Korrektur der Größe; herausfinden, wieviel Raum die Axis Beschriftung einnimmt ... 
                With .ChartObjects(anzDiagrams + 1)
                    .top = top
                    .height = 2 * height

                    Dim axleft As Double, axwidth As Double
                    If .Chart.HasAxis(Excel.XlAxisType.xlValue) = True Then
                        With CType(.Chart.Axes(Excel.XlAxisType.xlValue), Excel.Axis)
                            axleft = .Left
                            axwidth = .Width
                        End With
                        If left - axwidth < 1 Then
                            left = 1
                            width = width + left + 9
                        Else
                            left = left - axwidth
                            width = width + axwidth + 9
                        End If

                    End If

                    .left = left
                    .width = width


                End With


                'With .ChartObjects(anzDiagrams + 1)
                '    .top = top
                '    .left = left
                '    .height = 2 * height
                '    .width = width
                'End With

                '
                ' jetzt wird das zweite Diagramm gezeichnet - Gesamt-Bedarf des Projektes
                '
                If auswahl = 1 Then
                    diagramTitle = "Summe Ressourcen-Bedarf: " & gesamt_summe & zE & vbLf & pname
                Else
                    diagramTitle = "Summe Personalkosten: " & gesamt_summe & " T€" & vbLf & pname
                End If
                ReDim tdatenreihe(anzRollen - 1)
                ReDim Xdatenreihe(anzRollen - 1)
                For r = 1 To anzRollen
                    Xdatenreihe(r - 1) = CStr(ErgebnisListeR.Item(r))
                    tdatenreihe(r - 1) = hsum(r - 1)
                Next r

                With appInstance.Charts.Add
                    ' remove extra series

                    Do Until .SeriesCollection.Count = 0
                        .SeriesCollection(1).Delete()
                    Loop

                    With .SeriesCollection.NewSeries
                        .name = pname
                        .Values = tdatenreihe
                        .XValues = Xdatenreihe
                        .ChartType = Excel.XlChartType.xlPie
                        .HasDataLabels = True
                        .DataLabels.Position = Excel.XlDataLabelPosition.xlLabelPositionOutsideEnd
                    End With
                    For r = 1 To anzRollen
                        roleName = CStr(ErgebnisListeR.Item(r))
                        With .SeriesCollection(1).Points(r)
                            .Interior.color = RoleDefinitions.getRoledef(roleName).farbe
                            .DataLabel.Font.Size = 10
                        End With
                    Next r
                    '.HasLegend = False
                    .HasLegend = True
                    With .Legend
                        .Position = Excel.Constants.xlLeft
                        .Font.Size = awinSettings.fontsizeItems
                    End With

                    .HasTitle = True
                    .ChartTitle.text = diagramTitle
                    .ChartTitle.font.size = 10
                    .Location(Where:=XlChartLocation.xlLocationAsObject, Name:=appInstance.Worksheets(arrWsNames(3)).name)
                End With
                With .ChartObjects(anzDiagrams + 2)
                    .top = top
                    .left = left + width
                    .height = 10 * boxHeight
                    .width = 14 * boxWidth
                End With

            End If


        End With


        'Call awinScrollintoView()
        appInstance.EnableEvents = True
        appInstance.ScreenUpdating = True


    End Sub


    ''' <summary>
    ''' löscht die Einträge auf der Plantafel 
    ''' meist, um sie dann neu zeichnen zu können 
    ''' </summary>
    ''' <remarks></remarks>
    Public Sub awinClearPlanTafel()

        Dim formerEE As Boolean = appInstance.EnableEvents
        appInstance.EnableEvents = False


        ' jetzt müssen alle Shapes, die keine Charts sind, gelöscht werden ....

        For Each shp As Excel.Shape In CType(appInstance.ActiveSheet, Excel.Worksheet).Shapes

            Try
                If CBool(shp.HasChart) Then
                    ' do nothing, sollen ja erhalten bleiben 
                Else
                    shp.Delete()
                    
                End If
            Catch ex As Exception

            End Try


        Next shp

        projectboardShapes.clear()

        ' Änderung 26.7 weil Zahlen stehen blieben beim Neuladen einer neuen Konstellation
        If roentgenBlick.isOn Then
            With appInstance.Worksheets(arrWsNames(3))
                .range(.cells(2, 1), .cells(1000, 200)).clearcontents()
            End With
        End If

        ' jetzt werden für alle Projekte in Showprojekte die Verweise auf die Shapes gelöscht 
        For Each kvp As KeyValuePair(Of String, clsProjekt) In ShowProjekte.Liste
            kvp.Value.shpUID = ""
        Next

        ' jetzt wird die Zuordnung Projektname und Shape ID gelöscht ... 
        ShowProjekte.shpListe.Clear()

        appInstance.EnableEvents = formerEE

    End Sub

    Public Sub ClearPlanTafelfromOptArrows()

        Dim formerEE As Boolean = appInstance.EnableEvents
        appInstance.EnableEvents = False



        ' jetzt müssen alle Shapes, die Optmmierungs-Arrows sind, gelöscht werden ....


        For Each shp As Excel.Shape In CType(appInstance.ActiveSheet, Excel.Worksheet).Shapes
            With shp

                Try
                    If shp.AutoShapeType = MsoAutoShapeType.msoShapeRightArrow Or _
                    shp.AutoShapeType = MsoAutoShapeType.msoShapeLeftArrow Then
                        .Delete()
                    End If
                Catch ex As Exception

                End Try

            End With
        Next shp



        appInstance.EnableEvents = formerEE


    End Sub
    ' ''' <summary>
    ' ''' zeichnet die Plantafel mit den Projekten neu; 
    ' ''' versucht dabei immer die alte Position der Projekte zu übernehmen 
    ' ''' </summary>
    ' ''' <remarks></remarks>
    'Public Sub awinZeichnePlanTafel()

    '    Dim todoListe As New SortedList(Of Double, String)
    '    Dim key As Double
    '    Dim pname As String
    '    Dim zeile As Integer, lastZeile As Integer, curZeile As Integer, max As Integer
    '    Dim lastZeileOld As Integer
    '    Dim hproj As clsProjekt




    '    ' aufbauen der todoListe, so daß nachher die Projekte von oben nach unten gezeichnet werden können 
    '    For Each kvp As KeyValuePair(Of String, clsProjekt) In ShowProjekte.Liste

    '        With kvp.Value
    '            key = 10000 * .tfZeile + kvp.Value.Start
    '            todoListe.Add(key, .name)
    '        End With

    '    Next

    '    zeile = 2
    '    lastZeile = 0

    '    If ProjectBoardDefinitions.My.Settings.drawPhases = True Then
    '        ' dann sollen die Projekte im extended mode gezeichnet werden 
    '        ' jetzt erst mal die Konstellation "last" speichern
    '        Call awinStoreConstellation("Last")

    '        ' jetzt die todoListe abarbeiten
    '        For i = 1 To todoListe.Count
    '            pname = todoListe.ElementAt(i - 1).Value
    '            hproj = ShowProjekte.getProject(pname)

    '            If i = 1 Then
    '                curZeile = hproj.tfZeile
    '                lastZeileOld = hproj.tfZeile
    '                lastZeile = curZeile
    '                max = curZeile
    '            Else
    '                If lastZeileOld = hproj.tfZeile Then
    '                    curZeile = lastZeile
    '                Else
    '                    lastZeile = max
    '                    lastZeileOld = hproj.tfZeile
    '                End If

    '            End If

    '            hproj.tfZeile = curZeile
    '            lastZeile = curZeile
    '            'Call ZeichneProjektinPlanTafel2(pname, curZeile)
    '            Call ZeichneProjektinPlanTafel(pname, curZeile)
    '            curZeile = lastZeile + getNeededSpace(hproj)


    '            If curZeile > max Then
    '                max = curZeile
    '            End If


    '        Next

    '    Else


    '        Dim tryzeile As Integer

    '        For Each kvp As KeyValuePair(Of String, clsProjekt) In ShowProjekte.Liste
    '            pname = kvp.Key
    '            tryzeile = kvp.Value.tfZeile
    '            If tryzeile <= 1 Then
    '                tryzeile = -1
    '            End If
    '            Call ZeichneProjektinPlanTafel(pname, tryzeile) ' es wird versucht, an der alten Stelle zu zeichnen 
    '        Next


    '    End If





    'End Sub

    ''' <summary>
    ''' lädt die angegebene Projekt-Konstellation hinzu bzw. neu
    ''' funktioniert nur, wenn sowohl Konstellation als auch alle Projekte  im Hauptspeicher sind 
    ''' </summary>
    ''' <param name="constellationName">Name der Konstellation</param>
    ''' <param name="addProjects">gibt an, ob Projekte hinzugefügt werden sollen oder ob komplett neu gezeichnet werden soll</param>
    ''' <param name="storeLast">gibt an, ob die aktuelle Konstellation gespeichert werden soll</param>
    ''' <param name="updateProjektTafel">gibt an, ob die Projekt-Tafel neu gezeichnet werden soll oder ob die Konstellation nur im Showprojekte geladen werden soll</param> 
    ''' <remarks></remarks>
    Public Sub loadSessionConstellation(ByVal constellationName As String, ByVal addProjects As Boolean, ByVal storeLast As Boolean, _
                                        ByVal updateProjektTafel As Boolean)

        Dim activeConstellation As New clsConstellation
        Dim hproj As New clsProjekt
        Dim tfZeile As Integer
        Dim successMessage As String = ""
        Dim loadDateMessage As String = " * Das Datum kann nicht angepasst werden kann." & vbLf & _
                                        "   Das Projekt wurde bereits beauftragt."

        Dim projectDidNotExistYet As Boolean = True
        Dim firstFreeZeile As Integer = projectboardShapes.getMaxZeile + 1
        Dim anzahlNeueProjekte As Integer = 0


        ' prüfen, ob diese Constellation bereits existiert ..
        Try
            activeConstellation = projectConstellations.getConstellation(constellationName)
        Catch ex As Exception
            Call MsgBox(" Projekt-Konstellation " & constellationName & " existiert nicht ")
            Exit Sub
        End Try

        ' ggf die aktuelle Konstellation in "Last" speichern 

        If storeLast Then
            Call storeSessionConstellation(ShowProjekte, "Last")
        End If

        If Not addProjects And updateProjektTafel Then
            ShowProjekte.Clear()
            tfZeile = 2
        ElseIf addProjects And updateProjektTafel Then
            tfZeile = projectboardShapes.getMaxZeile + 1
        End If


        ' jetzt werden die Start-Values entsprechend gesetzt ..

        For Each kvp As KeyValuePair(Of String, clsConstellationItem) In activeConstellation.Liste

            If AlleProjekte.Containskey(kvp.Key) Then
                ' Projekt ist bereits im Hauptspeicher geladen

                hproj = AlleProjekte.getProject(kvp.Key)

                If ShowProjekte.contains(hproj.name) Then
                    With hproj
                        Call replaceProjectVariant(.name, .variantName, False, True, 0)
                        projectDidNotExistYet = False
                    End With
                ElseIf kvp.Value.show = True Then
                    ShowProjekte.Add(hproj)
                    projectDidNotExistYet = True
                    anzahlNeueProjekte = anzahlNeueProjekte + 1
                End If


                With hproj

                    ' Änderung THOMAS Start 
                    If .Status = ProjektStatus(0) Then
                        .startDate = kvp.Value.Start
                    ElseIf .startDate <> kvp.Value.Start Then
                        ' wenn das Datum nicht angepasst werden kann, weil das Projekt bereits beauftragt wurde  
                        successMessage = successMessage & vbLf & vbLf & loadDateMessage & vbLf & _
                                            "        " & hproj.name & ": " & kvp.Value.Start.ToShortDateString
                    End If
                    ' Änderung THOMAS Ende 

                    .StartOffset = 0

                    If projectDidNotExistYet And updateProjektTafel Then
                        If addProjects Then
                            .tfZeile = firstFreeZeile + anzahlNeueProjekte
                        Else
                            .tfZeile = kvp.Value.zeile
                        End If
                    End If


                End With



            Else

                Call MsgBox("Projekt " & kvp.Value.projectName & ", Variante: " & kvp.Value.variantName & vbLf & _
                             "ist nicht geladen!")

            End If


        Next

        If updateProjektTafel Then
            enableOnUpdate = False

            'appInstance.ScreenUpdating = False
            'Call diagramsVisible(False)
            Call awinClearPlanTafel()
            Call awinZeichnePlanTafel(False)
            Call awinNeuZeichnenDiagramme(2)
            'Call diagramsVisible(True)
            'appInstance.ScreenUpdating = True


            'Call MsgBox(constellationName & " wurde geladen ..." & vbLf & vbLf & successMessage)

            enableOnUpdate = True
        End If

        ' setzen der public variable, welche Konstellation denn jetzt gesetzt ist
        currentConstellation = constellationName


    End Sub


    Public Sub awinCalculateOptimization(ByVal diagrammTyp As String, ByRef myCollection As Collection, _
                                  ByRef OptimierungsErgebnis As SortedList(Of String, clsOptimizationObject))
        Dim referenceValue As Double, newReferenceValue As Double, currentValue As Double
        Dim bestValue As Double
        Dim startoffset As Integer
        Dim versatz As Integer
        Dim ErgebnisListe As New SortedList(Of Double, clsOptimizationObject)
        'Dim ErgebnisListe As New SortedDictionary(Of Double, clsOptimizationObject)
        Dim lokalesOptimum As clsOptimizationObject
        Dim hproj As clsProjekt
        Dim saveOffset As Integer, anzahlVersuche As Integer
        Dim NrLoops As Integer
        Dim NrArgExceptions As Integer
        Dim avgValue As Double



        If diagrammTyp = DiagrammTypen(0) Then ' Phase 
            Call MsgBox("Phasen Optimierung noch nicht implementiert")
        ElseIf diagrammTyp = DiagrammTypen(1) Or diagrammTyp = DiagrammTypen(2) Then

            With ShowProjekte
                If diagrammTyp = DiagrammTypen(1) Then
                    referenceValue = .getbadCostOfRole(myCollection)
                Else
                    referenceValue = .getDeviationfromAverage(myCollection, avgValue, diagrammTyp)
                End If

                newReferenceValue = -1
                NrLoops = 0
                NrArgExceptions = 0


                While newReferenceValue < referenceValue And NrLoops < 5 * ShowProjekte.Count
                    ' notwendig für den zweiten , ..n. durchlauf 
                    If newReferenceValue >= 0 Then
                        referenceValue = newReferenceValue
                    End If

                    For Each kvp As KeyValuePair(Of String, clsProjekt) In .Liste

                        If relevantForOptimization(kvp.Value) Then

                            bestValue = referenceValue  ' als Startwert, der hoffentlich unterboten wird .... 
                            startoffset = 0

                            For versatz = kvp.Value.earliestStart To kvp.Value.latestStart
                                If versatz <> 0 Then
                                    kvp.Value.StartOffset = versatz
                                    If diagrammTyp = DiagrammTypen(1) Then
                                        currentValue = .getbadCostOfRole(myCollection)
                                    Else
                                        currentValue = .getDeviationfromAverage(myCollection, avgValue, diagrammTyp)
                                    End If

                                    If currentValue < bestValue Then
                                        bestValue = currentValue
                                        startoffset = versatz
                                    End If
                                End If
                            Next versatz

                            ' zurücksetzen des StartOffsets im Projekt, weil hier ja erst verschiedene Konstellationen probiert werden  
                            kvp.Value.StartOffset = 0

                            If startoffset <> 0 Then ' es gab eine Verbesserung 
                                lokalesOptimum = New clsOptimizationObject
                                With lokalesOptimum
                                    .projectName = kvp.Key
                                    '.bestValue = bestValue
                                    .startOffset = startoffset
                                End With

                                Try
                                    ErgebnisListe.Add(bestValue, lokalesOptimum)
                                Catch ex As ArgumentException
                                    NrArgExceptions = NrArgExceptions + 1
                                    bestValue = bestValue + NrArgExceptions * 0.00000017
                                    Try
                                        ErgebnisListe.Add(bestValue, lokalesOptimum)
                                    Catch ex1 As ArgumentException
                                        NrArgExceptions = NrArgExceptions + 1
                                        bestValue = bestValue + NrArgExceptions * 0.00000017
                                        ErgebnisListe.Add(bestValue, lokalesOptimum)
                                    End Try
                                End Try
                            End If
                        End If

                    Next kvp
                    '
                    ' jetzt muss die Ergebnis Liste abgearbeitet werden ... 
                    '
                    anzahlVersuche = 0
                    newReferenceValue = referenceValue

                    For Each ergebnis As KeyValuePair(Of Double, clsOptimizationObject) In ErgebnisListe

                        hproj = ShowProjekte.getProject(ergebnis.Value.projectName)
                        saveOffset = hproj.StartOffset
                        hproj.StartOffset = ergebnis.Value.startOffset

                        If diagrammTyp = DiagrammTypen(1) Then
                            currentValue = .getbadCostOfRole(myCollection)
                        Else
                            currentValue = .getDeviationfromAverage(myCollection, avgValue, diagrammTyp)
                        End If

                        If currentValue < newReferenceValue Then
                            newReferenceValue = currentValue
                            anzahlVersuche = 0
                            ' hier müssen best, second, third gesetzt werden
                        Else
                            hproj.StartOffset = saveOffset
                            anzahlVersuche = anzahlVersuche + 1
                            If anzahlVersuche > 5 Or ergebnis.Value.startOffset = 0 Then
                                ' wenn startoffset = 0 , dann konnten keine Verbesserungen mehr erzielt werden , also Abbruch ...
                                Exit For
                            End If
                        End If
                        NrLoops = NrLoops + 1

                    Next

                    ErgebnisListe.Clear()

                End While
                ' hier wird Gold gesetzt, das heißt alle Offsets gemerkt, die für die Optmierung notwendig sind 
                ' anschließend werden alle startoffsets wieder auf 0 (=Ausgangswert) gesetzt 

                OptimierungsErgebnis.Clear()
                For Each kvp As KeyValuePair(Of String, clsProjekt) In .Liste
                    If kvp.Value.StartOffset <> 0 Then
                        lokalesOptimum = New clsOptimizationObject
                        With lokalesOptimum
                            .projectName = kvp.Value.name
                            '.bestValue = bestValue
                            .startOffset = kvp.Value.StartOffset
                        End With
                        OptimierungsErgebnis.Add(kvp.Value.name, lokalesOptimum)
                        'kvp.Value.StartOffset = 0
                    End If
                Next kvp

            End With

            'Case DiagrammTypen(2) ' Cost
            '    Call MsgBox("Kosten-Diagramm Optimierung noch nicht implementiert")
        Else
            Call MsgBox("Sonstige Optimierung noch nicht implementiert")
        End If

    End Sub

    Public Sub awinCalcOptimizationVarianten(ByVal diagrammTyp As String, ByRef myCollection As Collection, _
                                             ByVal worker As BackgroundWorker, ByVal e As DoWorkEventArgs)

        Dim anzahlVarianten As Integer
        Dim maxValue() As Integer
        Dim indexValue() As Integer
        Dim anzProjMitVar As Integer
        Dim PPointer As Integer
        Dim anzSchleifen As Integer = 0
        Dim firstValue As Double = 100000000000.0
        Dim secondValue As Double = 100000000000.0
        Dim thirdValue As Double = 100000000000.0
        Dim atleastOne As Boolean = False
        Dim anzKombinationen As Integer = 1
        Dim anzOptimierungen As Integer = 0

        Dim moreThanOne As New Collection
        Dim justOne As New Collection

        ' bestimme die Collection mit Projekten mit mehr als einer Variante
        For Each kvp As KeyValuePair(Of String, clsProjekt) In ShowProjekte.Liste
            anzahlVarianten = AlleProjekte.getVariantNames(kvp.Key, True).Count
            If anzahlVarianten = 1 Then
                justOne.Add(kvp.Key, kvp.Key)
            ElseIf anzahlVarianten > 1 Then
                moreThanOne.Add(kvp.Key, kvp.Key)
            End If
        Next


        If moreThanOne.Count = 0 Then
            e.Result = "es gibt keine Varianten .. demnach gibt es auch nichts zu optimieren !"
            worker.ReportProgress(0, e)
            Exit Sub
        Else
            ' speichern der letzten Konstellation
            Call storeSessionConstellation(ShowProjekte, autoSzenarioNamen(0))
        End If


        ' nimmt die Anzahl der Varianten auf
        anzProjMitVar = moreThanOne.Count
        ReDim maxValue(anzProjMitVar - 1)
        ReDim indexValue(anzProjMitVar - 1)

        ' jetzt wird bestimmt: 
        ' wievele Varianten hat das i.-te Element in morethanOne
        ' an welcher Stelle steht der Varianten-Zeiger Zeiger für das die bestimmt werden 
        Dim i As Integer = 0

        anzKombinationen = 1
        For Each pName As String In moreThanOne
            maxValue(i) = AlleProjekte.getVariantZahl(pName)
            ' in maxvalue steht 0, wenn es nur die Basis Variante gibt ..
            If maxValue(i) >= 0 Then
                anzKombinationen = anzKombinationen * (maxValue(i) + 1)
            End If

            indexValue(i) = 0
            i = i + 1
        Next

        PPointer = 0

        ' Start der Rekursion - und die Ausgangs-Konstellation als Vorgabe, als aktuelle "1. Varianten Optimum" behalten 
        firstValue = berechneOptimierungsWert(ShowProjekte, diagrammTyp, myCollection)
        Call storeSessionConstellation(ShowProjekte, autoSzenarioNamen(1))

        ' die anderen Szenarien sollen jetzt gelöscht werden 
        If projectConstellations.Contains(autoSzenarioNamen(2)) Then
            projectConstellations.Remove(autoSzenarioNamen(2))
        End If

        If projectConstellations.Contains(autoSzenarioNamen(3)) Then
            projectConstellations.Remove(autoSzenarioNamen(3))
        End If



        Call IterateOptimization(PPointer, anzProjMitVar, maxValue, indexValue, _
                                diagrammTyp, myCollection, anzKombinationen, anzSchleifen, anzOptimierungen, _
                                justOne, moreThanOne, _
                                firstValue, secondValue, thirdValue, _
                                worker, e)



        If anzOptimierungen > 0 Then
            ' wieder den alten Zustand herstellen 
            Call loadSessionConstellation(autoSzenarioNamen(0), False, False, False)
        Else
            ' es hat sich eh nichts geändert ... 
            'Call loadSessionConstellation(autoSzenarioNamen(0), False, False)
            e.Result = "in " & anzSchleifen.ToString & " Kombinationen" & vbLf & "konnte keine Verbesserung gefunden werden"
            worker.ReportProgress(0, e)

        End If


        ' erstelle alle Kombinationen der Varianten in der Variablen Current 


    End Sub




    ''' <summary>
    ''' rekursive Funktion, die die Kombinatorik der Varianten ermittelt 
    ''' </summary>
    ''' <param name="PPointer"></param>
    ''' <param name="anzProjMitVar"></param>
    ''' <param name="maxvalue"></param>
    ''' <param name="indexvalue"></param>
    ''' <param name="anzSchleifen"></param>
    ''' <remarks></remarks>
    Private Sub IterateOptimization(ByVal PPointer As Integer, ByVal anzProjMitVar As Integer, _
                                           ByVal maxvalue() As Integer, ByVal indexvalue() As Integer, _
                                           ByRef diagrammTyp As String, ByRef myCollection As Collection, _
                                           ByVal anzKombinationen As Integer, ByRef anzSchleifen As Integer, ByRef anzOptimierungen As Integer, _
                                           ByRef justOne As Collection, ByRef moreThanOne As Collection, _
                                           ByRef firstValue As Double, ByRef secondValue As Double, ByRef thirdValue As Double, _
                                           ByVal worker As BackgroundWorker, ByVal e As DoWorkEventArgs)

        'Dim currentSzenario As New clsProjekte
        Dim currentValue As Double
        Dim tmpConstellation As clsConstellation


        Dim hproj As clsProjekt

        If worker.CancellationPending Then
            e.Cancel = True
            e.Result = "Berichterstellung abgebrochen ..."
            Exit Sub
        End If


        If PPointer = anzProjMitVar - 1 Then

            indexvalue(anzProjMitVar - 1) = 0

            While indexvalue(anzProjMitVar - 1) <= maxvalue(anzProjMitVar - 1)


                ' jetzt die Aktion ausführen 

                Dim txtMSG As String = ""
                'currentSzenario = New clsProjekte
                For i = 1 To anzProjMitVar

                    If i = 1 Then
                        txtMSG = indexvalue(i - 1).ToString & ", "
                    ElseIf i = anzProjMitVar Then
                        txtMSG = txtMSG & indexvalue(i - 1).ToString
                    Else
                        txtMSG = txtMSG & indexvalue(i - 1).ToString & ", "
                    End If

                    hproj = AlleProjekte.getProject(CStr(moreThanOne.Item(i)), indexvalue(i - 1))
                    'currentSzenario.Add(hproj)

                    Call replaceProjectVariant(hproj.name, hproj.variantName, False, False, 0)

                Next

                'For Each pName As String In justOne
                '    hproj = AlleProjekte.getProject(calcProjektKey(pName, ""))
                '    currentSzenario.Add(hproj)
                'Next

                ' jetzt muss der Wert für current bestimmt werden 
                currentValue = berechneOptimierungsWert(ShowProjekte, diagrammTyp, myCollection)
                anzSchleifen = anzSchleifen + 1

                If currentValue < firstValue Then
                    thirdValue = secondValue
                    secondValue = firstValue
                    firstValue = currentValue

                    anzOptimierungen = anzOptimierungen + 1


                    If projectConstellations.Contains(autoSzenarioNamen(2)) Then
                        tmpConstellation = projectConstellations.getConstellation(autoSzenarioNamen(2))

                        If projectConstellations.Contains(autoSzenarioNamen(3)) Then
                            projectConstellations.Remove(autoSzenarioNamen(3))
                        End If

                        tmpConstellation.constellationName = autoSzenarioNamen(3)
                        projectConstellations.Add(tmpConstellation)

                        If projectConstellations.Contains(autoSzenarioNamen(1)) Then
                            tmpConstellation = projectConstellations.getConstellation(autoSzenarioNamen(1))

                            If projectConstellations.Contains(autoSzenarioNamen(2)) Then
                                projectConstellations.Remove(autoSzenarioNamen(2))
                            End If

                            tmpConstellation.constellationName = autoSzenarioNamen(2)
                            projectConstellations.Add(tmpConstellation)
                        End If


                    End If

                    Call storeSessionConstellation(ShowProjekte, autoSzenarioNamen(1))
                    'Call awinNeuZeichnenDiagramme(2)

                ElseIf currentValue < secondValue Then

                    anzOptimierungen = anzOptimierungen + 1

                    thirdValue = secondValue
                    secondValue = currentValue

                    If projectConstellations.Contains(autoSzenarioNamen(2)) Then
                        tmpConstellation = projectConstellations.getConstellation(autoSzenarioNamen(2))

                        If projectConstellations.Contains(autoSzenarioNamen(3)) Then
                            projectConstellations.Remove(autoSzenarioNamen(3))
                        End If

                        tmpConstellation.constellationName = autoSzenarioNamen(3)
                        projectConstellations.Add(tmpConstellation)

                    End If

                    Call storeSessionConstellation(ShowProjekte, autoSzenarioNamen(2))

                ElseIf currentValue < thirdValue Then

                    anzOptimierungen = anzOptimierungen + 1

                    thirdValue = currentValue
                    Call storeSessionConstellation(ShowProjekte, autoSzenarioNamen(3))
                End If

                e.Result = anzSchleifen.ToString & " / " & anzKombinationen.ToString & " Berechnungen; " & _
                            anzOptimierungen.ToString & " Optimierung(en"
                worker.ReportProgress(0, e)
                indexvalue(PPointer) = indexvalue(PPointer) + 1


            End While

            indexvalue(anzProjMitVar - 1) = 0


        Else

            For i = 0 To maxvalue(PPointer)
                indexvalue(PPointer) = i
                Call IterateOptimization(PPointer + 1, anzProjMitVar, maxvalue, indexvalue, _
                                        diagrammTyp, myCollection, anzKombinationen, anzSchleifen, anzOptimierungen, _
                                        justOne, moreThanOne, firstValue, secondValue, thirdValue, _
                                        worker, e)

                If worker.CancellationPending Then
                    e.Cancel = True
                    e.Result = "Berichterstellung abgebrochen ..."
                    Exit For
                End If

            Next


        End If




    End Sub

    ''' <summary>
    ''' bereichnet auf Basis der Freiheitsgrade der Projekte die beste Konstellation
    ''' </summary>
    ''' <param name="diagrammTyp"></param>
    ''' <param name="myCollection"></param>
    ''' <param name="OptimierungsErgebnis"></param>
    ''' <remarks></remarks>
    Public Sub awinCalcOptimizationFreiheitsgrade(ByVal diagrammTyp As String, ByRef myCollection As Collection, _
                                       ByRef OptimierungsErgebnis As SortedList(Of String, clsOptimizationObject))
        Dim currentValue As Double
        Dim bestValue As Double
        Dim startoffset As Integer
        Dim versatz As Integer
        Dim lokalesOptimum As New clsOptimizationObject
        Dim hproj As clsProjekt
        Dim NrArgExceptions As Integer
        Dim toDoListe As New Collection
        Dim NrLoops As Integer


        If myCollection.Count >= 1 Then


            If diagrammTyp = DiagrammTypen(0) Or diagrammTyp = DiagrammTypen(1) Or diagrammTyp = DiagrammTypen(2) Or diagrammTyp = DiagrammTypen(4) Then

                ' to do Liste aufbauen
                For Each kvp As KeyValuePair(Of String, clsProjekt) In ShowProjekte.Liste

                    If relevantForOptimization(kvp.Value) Then
                        toDoListe.Add(kvp.Key, kvp.Key)
                    End If
                Next kvp

                bestValue = berechneOptimierungsWert(ShowProjekte, diagrammTyp, myCollection)
                lokalesOptimum.bestValue = bestValue
                lokalesOptimum.projectName = " "
                OptimierungsErgebnis.Clear()

                NrLoops = 0
                NrArgExceptions = 0


                'While newReferenceValue < referenceValue And toDoListe.Count > 0
                Dim Abbruch As Boolean = False
                While toDoListe.Count > 0 And Not Abbruch

                    Dim i As Integer
                    Dim curProj As clsProjekt

                    For i = 1 To toDoListe.Count
                        curProj = ShowProjekte.getProject(CStr(toDoListe.Item(i)))

                        startoffset = 0

                        ' hier wird der beste Wert für das einzelne Projekt gesucht ....  

                        For versatz = curProj.earliestStart To curProj.latestStart
                            If versatz <> 0 Then
                                curProj.StartOffset = versatz
                                currentValue = berechneOptimierungsWert(ShowProjekte, diagrammTyp, myCollection)

                                If currentValue < bestValue Then
                                    bestValue = currentValue
                                    startoffset = versatz
                                End If
                            End If
                        Next versatz

                        ' zurücksetzen des StartOffsets im Projekt, weil hier ja erst verschiedene Konstellationen probiert werden  
                        curProj.StartOffset = 0

                        If startoffset <> 0 Then ' es gab eine Verbesserung 
                            'lokalesOptimum = New clsOptimizationObject
                            With lokalesOptimum
                                If bestValue < .bestValue Then
                                    .projectName = curProj.name
                                    .bestValue = bestValue
                                    .startOffset = startoffset
                                    ' Call awinVisualizeProject
                                End If
                            End With

                        End If

                    Next i
                    '
                    ' jetzt muss das Ergebnis abgearbeitet werden ... 
                    '
                    If lokalesOptimum.projectName <> " " Then

                        hproj = ShowProjekte.getProject(lokalesOptimum.projectName)
                        hproj.StartOffset = lokalesOptimum.startOffset
                        OptimierungsErgebnis.Add(lokalesOptimum.projectName, lokalesOptimum)
                        toDoListe.Remove(lokalesOptimum.projectName)
                        Call visualisiereTeilErgebnis(lokalesOptimum.projectName)
                    Else
                        Abbruch = True
                    End If

                    lokalesOptimum.projectName = " "
                    NrLoops = NrLoops + 1

                End While

            Else
                Call MsgBox("Optimierung noch nicht implementiert")
            End If
        Else
            Call MsgBox("Optimierung nicht implementiert")
        End If


    End Sub

    Public Function berechneOptimierungsWert(ByRef currentProjektListe As clsProjekte, ByRef DiagrammTyp As String, ByRef myCollection As Collection) As Double
        Dim value As Double
        Dim kennzahl1 As Double
        Dim kennzahl2 As Double
        Dim avgValue As Double

        If DiagrammTyp = DiagrammTypen(1) Then
            value = currentProjektListe.getbadCostOfRole(myCollection)

        ElseIf DiagrammTyp = DiagrammTypen(0) Then

            kennzahl1 = currentProjektListe.getAverage(myCollection, DiagrammTyp)
            kennzahl2 = currentProjektListe.getPhaseSchwellWerteInMonth(myCollection).Sum
            avgValue = System.Math.Max(kennzahl1, kennzahl2)
            value = currentProjektListe.getDeviationfromAverage(myCollection, avgValue, DiagrammTyp)

        ElseIf DiagrammTyp = DiagrammTypen(2) Then
            avgValue = currentProjektListe.getAverage(myCollection, DiagrammTyp)
            value = currentProjektListe.getDeviationfromAverage(myCollection, avgValue, DiagrammTyp)

        ElseIf DiagrammTyp = DiagrammTypen(4) Then
            ' da der Optimierungs-Algorithmus die kleinste Zahl sucht , muss mit -1 multipliziert werden, 
            ' damit tatsächlich der größte Ertrag heraus kommt 
            value = currentProjektListe.getErgebniskennzahl * (-1)

        ElseIf DiagrammTyp = DiagrammTypen(5) Then
            'Throw New ArgumentException("Optimierung ist für diesen Diagramm-Typ nicht implementiert")
            ' tk: das folgende kann aktiviert werden, sobald 
            kennzahl1 = currentProjektListe.getAverage(myCollection, DiagrammTyp)
            kennzahl2 = currentProjektListe.getMilestoneSchwellWerteInMonth(myCollection).Sum
            avgValue = System.Math.Max(kennzahl1, kennzahl2)
            value = currentProjektListe.getDeviationfromAverage(myCollection, avgValue, DiagrammTyp)
        Else
            Throw New ArgumentException("Optimierung ist für diesen Diagramm-Typ nicht implementiert")
        End If

        berechneOptimierungsWert = value

    End Function
    Public Function relevantForOptimization(ByRef project As clsProjekt) As Boolean
        Dim relevant As Boolean = False
        Dim bereichsAnfang As Integer, bereichsEnde As Integer
        ' hier können dann auch weitere Bedingungen abgefragt werden, ob bspweise das Projekt denn überhaupt Freiheitsgrade besitzt  


        With project

            If .Status = ProjektStatus(0) Then

                ' nur dann darf das Projekt noch verschoben werden ...

                bereichsAnfang = .Start + .earliestStart
                bereichsEnde = .Start + .anzahlRasterElemente - 1 + .latestStart

                If project.StartOffset = 0 And (project.earliestStart < 0 Or project.latestStart > 0) _
                                               And istBereichInTimezone(bereichsAnfang, bereichsEnde) Then
                    relevant = True
                Else
                    relevant = False
                End If
            Else
                relevant = False
            End If

        End With

        relevantForOptimization = relevant

    End Function

    ''' <summary>
    ''' zeichnet für alle selektierten Projekte die Phasen, die in Namelist angegeben sind;  
    ''' wenn namelist leer ist, werden alle Phasen des Projektes angezeigt
    ''' </summary>
    ''' <param name="nameList">enthält die Namen plus die Breadcrumbs der Phasen, die gezeichnet werden sollen; alle, wenn leer</param>
    ''' <param name="numberIt">gibt an, ob di ePhasen nummeriert werden sollen</param>
    ''' <param name="deleteOtherShapes">gibt an, ob die anderen Phasen-Shapes gelöscht werden sollen</param>
    ''' <remarks></remarks>
    Public Sub awinZeichnePhasen(ByVal nameList As Collection, ByVal numberIt As Boolean, ByVal deleteOtherShapes As Boolean)

        'Dim request As New Request(awinSettings.databaseName)
        Dim singleShp As Excel.Shape
        Dim hproj As New clsProjekt
        Dim vglName As String = " "
        Dim pName As String
        Dim ok As Boolean = True
        Dim msNumber As Integer = 1

        Dim awinSelection As Excel.ShapeRange

        Dim formerEE As Boolean = appInstance.EnableEvents
        Dim formerSU As Boolean = appInstance.ScreenUpdating
        appInstance.EnableEvents = False
        appInstance.ScreenUpdating = True

        enableOnUpdate = False

        Try
            awinSelection = CType(appInstance.ActiveWindow.Selection.ShapeRange, Excel.ShapeRange)
        Catch ex As Exception
            awinSelection = Nothing
        End Try


        If Not awinSelection Is Nothing Then

            Dim anzSelect As Integer = awinSelection.Count

            ' jetzt die Aktion durchführen ...
          
            For Each singleShp In awinSelection
                ok = True
                With singleShp
                    If isProjectType(kindOfShape(singleShp)) Then


                        Try
                            hproj = ShowProjekte.getProject(singleShp.Name)
                        Catch ex As Exception
                            ok = False
                        End Try

                        If ok Then

                            If deleteOtherShapes Then
                                Call awinDeleteProjectChildShapes(singleShp, 3)
                            End If

                            Try
                                pName = hproj.name
                                Call zeichnePhasenInProjekt(hproj, nameList, False, msNumber)

                            Catch ex As Exception

                            End Try


                        End If

                    End If
                End With
            Next

            If msNumber = 1 Then
                If nameList.Count > 1 Then
                    Call MsgBox("Auswahl enthält diese Phasen nicht")
                Else
                    Call MsgBox("Auswahl enthält diese Phase nicht:  " & nameList.Item(1))
                End If
            End If

        Else
            ' tue es für alle Projekte in Showprojekte 


            Dim todoListe As New SortedList(Of Long, clsProjekt)
            Dim key As Long

            For Each kvp As KeyValuePair(Of String, clsProjekt) In ShowProjekte.Liste

                key = 10000 * kvp.Value.tfZeile + kvp.Value.tfspalte
                todoListe.Add(key, kvp.Value)

            Next


            For Each kvp As KeyValuePair(Of Long, clsProjekt) In todoListe

                If deleteOtherShapes Then
                    singleShp = ShowProjekte.getShape(kvp.Value.name)
                    Call awinDeleteProjectChildShapes(singleShp, 3)
                End If

                Call zeichnePhasenInProjekt(kvp.Value, nameList, False, msNumber, showRangeLeft, showRangeRight)

            Next


            If msNumber = 1 Then
                If nameList.Count > 1 Then
                    Call MsgBox("im gewählten Zeitraum gibt es diese Phasen nicht")
                Else
                    Call MsgBox("im gewählten Zeitraum gibt es diese Phase nicht: " & nameList.Item(1))
                End If
            End If


        End If


        'ur: 17.7.2015: für PlanElemente visualisieren für Einzelprojekt-Info sollte nach zeichnen der Phasen nicht deselektiert werden
        '' ''Call awinDeSelect()


        enableOnUpdate = True
        appInstance.EnableEvents = formerEE
        appInstance.ScreenUpdating = formerSU



    End Sub

    ''' <summary>
    ''' zeichnet die Ressourcen- bzw. Kostenbedarfe in die Projekt-Tafel 
    ''' </summary>
    ''' <param name="nameList"></param>
    ''' <param name="prcTyp"></param>
    ''' <remarks></remarks>
    Public Sub awinZeichneBedarfe(ByVal nameList As Collection, ByVal prcTyp As String)

        Dim tmpName As String = ""

        If nameList.Count < 1 Then
            tmpName = ""
        ElseIf nameList.Count = 1 Then
            tmpName = CStr(nameList.Item(1))
        ElseIf nameList.Count > 1 Then
            tmpName = "Collection"
            
        End If

        With roentgenBlick
            If .isOn Then
                Call awinNoshowProjectNeeds()
            End If
            .isOn = True
            .name = tmpName
            .myCollection = nameList
            .type = prcTyp
            Call awinShowProjectNeeds1(nameList, prcTyp)
        End With
    End Sub

    ''' <summary>
    ''' zeichnet für interaktiven wie Report Modus die Milestones 
    ''' 0: grau, 1: grün, 2: gelb, 3:rot, 4: alle
    ''' </summary>
    ''' <param name="farbTyp">welcher Typus soll gezeichnet werden </param>
    ''' <remarks></remarks>
    Public Sub awinZeichneMilestones(ByVal nameList As Collection, ByVal farbTyp As Integer, ByVal numberIt As Boolean, ByVal deleteOtherShapes As Boolean)

        'Dim request As New Request(awinSettings.databaseName)
        Dim singleShp As Excel.Shape
        Dim hproj As New clsProjekt
        Dim vglName As String = " "
        Dim pName As String
        Dim ok As Boolean = True
        Dim msNumber As Integer = 1

        Dim awinSelection As Excel.ShapeRange

        Dim formerEE As Boolean = appInstance.EnableEvents
        Dim formerSU As Boolean = appInstance.ScreenUpdating
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
                ok = True
                With singleShp

                    If isProjectType(kindOfShape(singleShp)) Then

                        Try
                            hproj = ShowProjekte.getProject(singleShp.Name)
                        Catch ex As Exception
                            ok = False
                        End Try

                        If ok Then

                            If deleteOtherShapes Then
                                Call awinDeleteProjectChildShapes(singleShp, 1)
                            End If

                            Try
                                pName = hproj.name
                                Call zeichneMilestonesInProjekt(hproj, nameList, farbTyp, 0, 0, False, msNumber, False)
                            Catch ex As Exception
                                Dim a As Integer = 0
                            End Try


                        End If

                    End If
                End With
            Next

            If msNumber = 1 Then
                If nameList.Count > 1 Then
                    Call MsgBox("Auswahl enthält  diese Meilensteine nicht")
                ElseIf nameList.Count = 1 Then
                    Call MsgBox("Auswahl enthält keinen Meilenstein " & nameList.Item(1))

                End If
            End If

        Else


            If ShowProjekte.Count > 0 Then

                ' tue es für alle Projekte in Showprojekte 


                Dim todoListe As New SortedList(Of Long, clsProjekt)
                Dim key As Long

                For Each kvp As KeyValuePair(Of String, clsProjekt) In ShowProjekte.Liste

                    key = 10000 * kvp.Value.tfZeile + kvp.Value.tfspalte
                    todoListe.Add(key, kvp.Value)

                Next


                For Each kvp As KeyValuePair(Of Long, clsProjekt) In todoListe

                    If deleteOtherShapes Then
                        singleShp = ShowProjekte.getShape(kvp.Value.name)
                        Call awinDeleteProjectChildShapes(singleShp, 1)
                    End If

                    Call zeichneMilestonesInProjekt(kvp.Value, nameList, farbTyp, showRangeLeft, showRangeRight, numberIt, msNumber, False)

                Next


                If msNumber = 1 Then
                    If nameList.Count > 1 Then
                        Call MsgBox("im gewählten Zeitraum gibt es diese Meilensteine nicht")
                    ElseIf nameList.Count = 1 Then
                        Call MsgBox("im gewählten Zeitraum gibt es keinen Meilenstein " & nameList.Item(1))
                    End If
                End If

            Else
                Call MsgBox("Es sind keine Projekte geladen!")
            End If


        End If

        'ur: 17.7.2015: für PlanElemente visualisieren sollte nach zeichnen der Meilensteine nicht deselektiert werden
        '' ''Call awinDeSelect()

        enableOnUpdate = True
        appInstance.EnableEvents = formerEE
        appInstance.ScreenUpdating = formerSU



    End Sub


    ''' <summary>
    ''' bringt alle charts in den Vordergrund, so daß sie nicht von einem neu gezeichneten Projekt überdeckt werden 
    ''' </summary>
    ''' <remarks></remarks>
    Public Sub bringChartsToFront(ByVal projectShape As Excel.Shape)

        Dim worksheetShapes As Excel.Shapes
        Dim chtobj As Excel.ChartObject

        ' sicherstellen, dass projectshape auch etwas enthält ... 
        If IsNothing(projectShape) Then
            Exit Sub
        End If

        Try

            worksheetShapes = CType(appInstance.Worksheets(arrWsNames(3)), Excel.Worksheet).Shapes

        Catch ex As Exception
            Throw New Exception("in bringChartstoFront : keine Shapes Zuordnung möglich ")
        End Try


        For Each chtobj In CType(CType(appInstance.Worksheets(arrWsNames(3)), Excel.Worksheet).ChartObjects, Excel.ChartObjects)

            Try
            With chtobj
                    If ((projectShape.Top >= .Top And projectShape.Top <= .Top + .Height) Or _
                        (.Top >= projectShape.Top And .Top <= projectShape.Top + projectShape.Height)) And _
                        ((projectShape.Left >= .Left And projectShape.Left <= .Left + .Width) Or _
                        (.Left >= projectShape.Left And .Left <= projectShape.Left + projectShape.Width)) Then

                        CType(worksheetShapes.Item(chtobj.Name), Excel.Shape).ZOrder(MsoZOrderCmd.msoBringToFront)

                    End If
                End With
            Catch ex As Exception

            End Try


        Next


    End Sub

    ''' <summary>
    ''' zeichnet die Plantafel mit den Projekten neu; 
    ''' zeichnet bei fromScratch = true: zuerst in Reihenfolge der Business Units, 
    ''' dann sortiert nach Anfangsdatum, dann sortiert nach Projektdauer
    ''' im fall fromScratch = false: versucht dabei immer die alte Position der Projekte zu übernehmen 
    ''' </summary>
    ''' <remarks></remarks>
    Public Sub awinZeichnePlanTafel(ByVal fromScratch As Boolean)

        Dim todoListe As New SortedList(Of Double, String)
        Dim key As Double
        Dim pname As String

        Dim lastZeileOld As Integer
        Dim hproj As clsProjekt
        Dim positionsKennzahl As Double

        Dim notOK As Boolean = True




        If fromScratch Then
            Dim zeile As Integer
            Dim lastBU As String = ""

            For Each kvp As KeyValuePair(Of String, clsProjekt) In ShowProjekte.Liste

                notOK = True

                With kvp.Value

                    positionsKennzahl = calcKennziffer(kvp.Value)

                    Do While notOK
                        Try
                            todoListe.Add(positionsKennzahl, .name)
                            notOK = False
                        Catch ex As Exception
                            positionsKennzahl = positionsKennzahl + 0.01
                        End Try
                    Loop


                End With

            Next

            zeile = 2
            Dim i As Integer

            For i = 1 To todoListe.Count

                pname = todoListe.ElementAt(i - 1).Value

                Try
                    hproj = ShowProjekte.getProject(pname)

                    If i = 1 Then
                        lastBU = hproj.businessUnit
                    ElseIf lastBU <> hproj.businessUnit Then
                        lastBU = hproj.businessUnit
                        zeile = zeile + 1
                    End If

                    hproj.tfZeile = zeile

                    Dim tmpCollection As New Collection
                    Call ZeichneProjektinPlanTafel(tmpCollection, pname, zeile, tmpCollection, tmpCollection)

                    zeile = zeile + hproj.calcNeededLines(tmpCollection, awinSettings.drawphases, False)

                Catch ex As Exception

                End Try

            Next


        Else

            Dim zeile As Integer, lastzeile As Integer, curzeile As Integer, max As Integer
            ' so wurde es bisher gemacht ... bis zum 17.1.15
            ' aufbauen der todoListe, so daß nachher die Projekte von oben nach unten gezeichnet werden können 
            For Each kvp As KeyValuePair(Of String, clsProjekt) In ShowProjekte.Liste

                With kvp.Value
                    key = 10000 * .tfZeile + kvp.Value.Start
                    todoListe.Add(key, .name)
                End With

            Next

            zeile = 2
            lastZeile = 0


            'If ProjectBoardDefinitions.My.Settings.drawPhases = True Then
            ' dann sollen die Projekte im extended mode gezeichnet werden 
            ' jetzt erst mal die Konstellation "last" speichern
            ' 3.11.14 Auskommentiert: Zeichnen sollte nichts zu tun haben mit dem Verwalten von Konstellationen 
            ' Call storeSessionConstellation(ShowProjekte, "Last")

            ' jetzt die todoListe abarbeiten
            Dim i As Integer
            For i = 1 To todoListe.Count
                pname = todoListe.ElementAt(i - 1).Value

                Try
                    hproj = ShowProjekte.getProject(pname)

                    If i = 1 Then
                        curZeile = hproj.tfZeile
                        lastZeileOld = hproj.tfZeile
                        lastZeile = curZeile
                        max = curZeile
                    Else
                        If lastZeileOld = hproj.tfZeile Then
                            curZeile = lastZeile
                        Else
                            lastZeile = max
                            lastZeileOld = hproj.tfZeile
                        End If

                    End If

                    ' Änderung 9.10.14, damit die Spaces in einer 
                    'If hproj.tfZeile >= curZeile + 1 Then
                    '    curZeile = curZeile + 1
                    'End If
                    ' Ende Änderung
                    hproj.tfZeile = curZeile
                    lastZeile = curZeile
                    'Call ZeichneProjektinPlanTafel2(pname, curZeile)
                    ' wenn bestimmte Projekte beim Suchen nach einem Platz nicht berücksichtigt werden sollen,
                    ' dann müssen sie in einer Collection an ZeichneProjektinPlanTafel übergeben werden 
                    Dim tmpCollection As New Collection
                    Call ZeichneProjektinPlanTafel(tmpCollection, pname, curZeile, tmpCollection, tmpCollection)
                    curzeile = lastzeile + hproj.calcNeededLines(tmpCollection, awinSettings.drawphases, False)


                    If curZeile > max Then
                        max = curZeile
                    End If
                Catch ex As Exception

                End Try



            Next
        End If




    End Sub


    ''' <summary>
    ''' zeichnet das Projekt "pname" in die Plantafel; 
    ''' wenn es bereits vorhanden ist: keine Aktion  
    ''' noCollection ist eine Collection von Projekt-Namen, die beim Suchen nach einem Platz 
    ''' auf der Projekt-Tafel nicht berücksichtigt werden soll
    ''' ist insbesondere wichtig, wenn mehrere Projekte selektiert wurden und verschoben werden 
    ''' </summary>
    ''' <param name="pname"></param>
    ''' <remarks></remarks>
    Public Sub ZeichneProjektinPlanTafel(ByVal noCollection As Collection, ByVal pname As String, ByVal tryzeile As Integer, _
                                         ByVal drawPhaseList As Collection, ByVal drawMilestoneList As Collection)


        Dim drawphases As Boolean = awinSettings.drawphases
        Dim phasenNameID As String
        Dim phaseShapeName As String
        Dim msShapeName As String

        Dim start As Integer
        Dim laenge As Integer
        Dim status As String
        Dim pMarge As Double
        Dim pcolor As Object, schriftfarbe As Object
        Dim schriftgroesse As Integer
        Dim zeile As Integer
        Dim hproj As clsProjekt
        Dim top As Double, left As Double, width As Double, height As Double
        Dim projectShape As Excel.Shape
        Dim phaseShape As Excel.Shape, milestoneShape As Excel.Shape
        Dim shpUID As String
        'Dim tmpshapes As Excel.Shapes = appInstance.ActiveSheet.shapes
        Dim worksheetShapes As Excel.Shapes
        Dim heute As Date = Date.Now
        Dim tmpShapeRange As Excel.ShapeRange
        Dim vorlagenShape As xlNS.Shape

        Dim shpExists As Boolean
        Dim oldAlternativeText As String = ""


        Try

            worksheetShapes = CType(appInstance.Worksheets(arrWsNames(3)), Excel.Worksheet).Shapes

        Catch ex As Exception
            Throw New Exception("in ZeichneProjektinPlanTafel : keine Shapes Zuordnung möglich ")
        End Try

        Try
            hproj = ShowProjekte.getProject(pname)
            With hproj
                laenge = .anzahlRasterElemente
                shpUID = .shpUID
                start = .Start + .StartOffset
                pcolor = .farbe
                schriftfarbe = .Schriftfarbe
                schriftgroesse = .Schrift
                status = .Status
                pMarge = .ProjectMarge
            End With
        Catch ex As Exception
            Throw New ArgumentException("in zeichneProjektinBoard - Projektname existiert nicht: " & pname)
        End Try


        ' prüfen, ob das Shape bereits existiert ...
        If shpUID <> "" Then
            Try
                projectShape = worksheetShapes.Item(pname)
                shpExists = True
                ' merken, weil bei Variante erzeugen der Alternative Text nicht geändert werden soll 
                oldAlternativeText = projectShape.AlternativeText
            Catch ex As Exception
                shpExists = False
                projectShape = Nothing
            End Try
        Else
            shpExists = False
            projectShape = Nothing
        End If



        '
        ' ist dort überhaupt Platz ? wenn nicht, dann Zeile mit freiem Platz suchen ...
        If tryzeile < 2 Then
            tryzeile = projectboardShapes.getMaxZeile
        End If


        zeile = findeMagicBoardPosition(noCollection, pname, tryzeile, start, laenge)


        Dim formerEE As Boolean = appInstance.EnableEvents
        enableOnUpdate = False
        appInstance.EnableEvents = False



        If shpExists Then

            If drawphases Then

                ' ungroup Shape, damit die einzelnen Phasen- bzw Milestone Shapes im Zugriff sind 
                Try
                    tmpShapeRange = projectShape.Ungroup
                Catch ex As Exception
                    tmpShapeRange = Nothing
                End Try

                Dim cphase As clsPhase

                For i = 1 To hproj.CountPhases
                    cphase = hproj.getPhase(i)
                    phasenNameID = cphase.nameID
                    phaseShapeName = projectboardShapes.calcPhaseShapeName(pname, phasenNameID) & "#" & i.ToString
                    'phaseShapeName = pname & "#" & phasenName & "#" & i.ToString

                    Try
                        phaseShape = worksheetShapes.Item(phaseShapeName)
                        Call defineShapeAppearance(hproj, phaseShape, i)
                    Catch ex As Exception

                    End Try

                    For r = 1 To cphase.countMilestones

                        Dim cMilestone As clsMeilenstein
                        Dim cBewertung As clsBewertung

                        cMilestone = cphase.getMilestone(r)
                        cBewertung = cMilestone.getBewertung(1)

                        msShapeName = projectboardShapes.calcMilestoneShapeName(hproj.name, cMilestone.nameID)

                        ' existiert das schon ? 
                        Try
                            milestoneShape = worksheetShapes.Item(msShapeName)
                            Call defineResultAppearance(hproj, 0, milestoneShape, cBewertung)
                        Catch ex As Exception

                        End Try


                    Next


                Next

                ' Gruppieren des Shapes 
                projectShape = tmpShapeRange.Group
                projectShape.Name = hproj.name



            Else

                Call defineShapeAppearance(hproj, projectShape)

            End If


        Else

            ' ///////////////
            ' Shape existiert noch nicht 
            ' ///////////////

            ' hier wird der vorher bestimmte Wert gesetzt, wo das Shape gezeichnet werden kann 
            hproj.tfZeile = zeile

            If drawphases And (hproj.CountPhases > 1) Then
                ' stelle das Projekt im Extended Mode dar  
                'Dim shapeGroupListe() As Object
                Dim shapeGroupListe() As String
                Dim arrayOfMSNames() As String
                Dim msShapeNames As New Collection
                Dim anzGroupElemente As Integer = 0
                Dim projectShapesCollection As New Collection



                'oldShape = Nothing
                phaseShape = Nothing

                Dim zeilenOffset As Integer = 0
                Dim lastEndDate As Date = StartofCalendar.AddDays(-1)

                For i = 1 To hproj.CountPhases

                    With hproj.getPhase(i)

                        phasenNameID = .nameID

                    End With



                    Try
                        zeilenOffset = 0
                        Call hproj.calculateShapeCoord(i, zeilenOffset, top, left, width, height)

                        If i = 1 Then

                            If awinSettings.drawProjectLine Then

                                phaseShape = worksheetShapes.AddConnector(MsoConnectorType.msoConnectorStraight, CSng(left), CSng(top), _
                                                                            CSng(left + width), CSng(top))
                            Else

                                phaseShape = worksheetShapes.AddShape(Type:=Microsoft.Office.Core.MsoAutoShapeType.msoShapeRoundedRectangle, _
                                                        Left:=CSng(left), Top:=CSng(top), Width:=CSng(width), Height:=CSng(height))
                            End If

                        Else
                            vorlagenShape = PhaseDefinitions.getShape(elemNameOfElemID(phasenNameID))

                            phaseShape = worksheetShapes.AddShape(Type:=vorlagenShape.AutoShapeType, _
                                Left:=CSng(left), Top:=CSng(top), Width:=CSng(width), Height:=CSng(height))
                            vorlagenShape.PickUp()
                            phaseShape.Apply()
                        End If



                    Catch ex As Exception
                        Throw New Exception("in zeichneProjektinPlantafel2 : keine Shape-Erstellung möglich ...  ")
                    End Try

                    phaseShapeName = projectboardShapes.calcPhaseShapeName(pname, phasenNameID) & "#" & i.ToString
                    'phaseShapeName = pname & "#" & phasenName & "#" & i.ToString
                    With phaseShape
                        .Name = phaseShapeName
                        .Title = phasenNameID
                        .AlternativeText = CInt(PTshty.phaseE).ToString
                    End With

                    If i = 1 And awinSettings.drawProjectLine Then
                        Call defineShapeAppearance(hproj, phaseShape)
                    Else
                        Call defineShapeAppearance(hproj, phaseShape, i)
                    End If


                    ' jetzt der Liste der ProjectboardShapes hinzufügen
                    projectboardShapes.add(phaseShape)

                    Try
                        projectShapesCollection.Add(phaseShapeName, Key:=phaseShapeName)
                    Catch ex As Exception

                    End Try


                    ' jetzt müssen alle Meilensteine dieser Phase gezeichnet werden 

                    With CType(hproj.getPhase(i), clsPhase)
                        Dim msName As String
                        Dim msShape As Excel.Shape

                        For r = 1 To .countMilestones

                            Dim cMilestone As clsMeilenstein
                            Dim cBewertung As clsBewertung

                            cMilestone = .getMilestone(r)
                            cBewertung = cMilestone.getBewertung(1)

                            vorlagenShape = MilestoneDefinitions.getShape(cMilestone.name)
                            Dim factorB2H As Double = vorlagenShape.Width / vorlagenShape.Height

                            hproj.calculateMilestoneCoord(cMilestone.getDate, zeilenOffset, factorB2H, top, left, width, height)

                            msName = projectboardShapes.calcMilestoneShapeName(hproj.name, cMilestone.nameID)
                            'msName = hproj.name & "#" & .name & "#M" & r.ToString
                            ' existiert das schon ? 
                            Try
                                msShape = worksheetShapes.Item(msName)
                            Catch ex As Exception
                                msShape = Nothing
                            End Try

                            If msShape Is Nothing Then


                                'msShape = worksheetShapes.AddShape(Type:=Microsoft.Office.Core.MsoAutoShapeType.msoShapeDiamond, _
                                '                                Left:=CSng(left), Top:=CSng(top), Width:=CSng(width), Height:=CSng(height))
                                msShape = worksheetShapes.AddShape(Type:=vorlagenShape.AutoShapeType, _
                                                                Left:=CSng(left), Top:=CSng(top), Width:=CSng(width), Height:=CSng(height))
                                vorlagenShape.PickUp()
                                msShape.Apply()

                                With msShape
                                    .Name = msName
                                    .Title = cMilestone.nameID
                                    .AlternativeText = CInt(PTshty.milestoneE).ToString
                                End With


                                ' tk 24.3.2015 um nachher die Milestone Shapes nach vorne zu holen
                                If Not msShapeNames.Contains(msName) Then
                                    msShapeNames.Add(msName, msName)
                                End If

                                msShape.Rotation = vorlagenShape.Rotation

                                Call defineResultAppearance(hproj, 0, msShape, cBewertung)

                                ' jetzt der Liste der ProjectboardShapes hinzufügen
                                projectboardShapes.add(msShape)

                            Else
                                ' Koordinaten anpassen 
                                msShape.Top = CSng(top)
                            End If

                            Try
                                projectShapesCollection.Add(msName, Key:=msName)
                            Catch ex As Exception

                            End Try

                        Next


                    End With

                Next

                ' Änderung tk 24.3.2015
                ' jetzt müssen ggf die Meilensteine noch nach vorne gebracht werden ...
                Dim anzElements As Integer
                anzElements = msShapeNames.Count

                If anzElements > 0 Then

                    ReDim arrayOfMSNames(anzElements - 1)
                    For ix = 1 To anzElements
                        arrayOfMSNames(ix - 1) = CStr(msShapeNames.Item(ix))
                    Next

                    Try
                        CType(worksheetShapes.Range(arrayOfMSNames), Excel.ShapeRange).ZOrder(MsoZOrderCmd.msoBringToFront)
                    Catch ex As Exception

                    End Try

                End If


                ' hier werden die Shapes gruppiert
                anzGroupElemente = projectShapesCollection.Count

                If anzGroupElemente > 1 Then
                    ' es macht nur Sinn zu gruppieren, wenn es mehr als 1 Element ist ....

                    ReDim shapeGroupListe(anzGroupElemente - 1)
                    For i = 1 To anzGroupElemente
                        shapeGroupListe(i - 1) = CStr(projectShapesCollection.Item(i))
                    Next

                    Dim ShapeGroup As Excel.ShapeRange
                    ShapeGroup = worksheetShapes.Range(shapeGroupListe)
                    projectShape = ShapeGroup.Group()

                Else
                    ' in diesem Fall besteht das Projekt nur aus einer einzigen Phase
                    projectShape = phaseShape

                End If
                projectShape.Name = pname


            Else
                ' stelle das Projekt im Einzeilen Mode dar

                With hproj
                    .CalculateShapeCoord(top, left, width, height)
                    .tfZeile = zeile
                End With

                If awinSettings.drawProjectLine Then

                    projectShape = worksheetShapes.AddConnector(MsoConnectorType.msoConnectorStraight, CSng(left), CSng(top), _
                                                                CSng(left + width), CSng(top))

                Else
                    projectShape = worksheetShapes.AddShape(Type:=Microsoft.Office.Core.MsoAutoShapeType.msoShapeRoundedRectangle, _
                        Left:=CSng(left), Top:=CSng(top), Width:=CSng(width), Height:=CSng(height))

                End If

                projectShape.Name = pname
                Call defineShapeAppearance(hproj, projectShape)

            End If


        End If

        With projectShape
            If shpExists Then
                .AlternativeText = oldAlternativeText
            Else
                If drawphases Then
                    .AlternativeText = CInt(PTshty.projektE).ToString
                Else
                    If awinSettings.drawProjectLine Then
                        .AlternativeText = CInt(PTshty.projektL).ToString
                    Else
                        .AlternativeText = CInt(PTshty.projektN).ToString
                    End If
                End If
            End If


            hproj.shpUID = .ID.ToString
            hproj.tfZeile = calcYCoordToZeile(projectShape.Top)
        End With

        ' jetzt der Liste der ProjectboardShapes hinzufügen
        projectboardShapes.add(projectShape)

        ' jetzt muss das neue Shape in der ShowProjekte.ShapeListe eingetragen werden ..
        ShowProjekte.AddShape(pname, shpUID:=projectShape.ID.ToString)

        ' jetzt müssen ggf die noch zu zeichnenden Meilensteine und Phasen eingezeichnet werden  

        Dim msNumber As Integer = 0
        If drawPhaseList.Count > 0 And Not drawphases Then
            Call zeichnePhasenInProjekt(hproj, drawPhaseList, False, msNumber)
        End If

        msNumber = 0
        If drawMilestoneList.Count > 0 And Not drawphases Then
            Call zeichneMilestonesInProjekt(hproj, drawMilestoneList, 4, 0, 0, False, msNumber, False)
        End If

        ' zu guter Letzt muss der Projekt-Name gezeichnet werden 
        If awinSettings.drawProjectLine Then
            Call zeichneNameInProjekt(hproj)
        End If


        If roentgenBlick.isOn Then
            With roentgenBlick
                Call awinShowNeedsofProject1(mycollection:=.myCollection, type:=.type, projektname:=pname)
            End With
        End If

        If drawPhaseList.Count = 0 And drawMilestoneList.Count = 0 Then
            ' jetzt müssen die Charts, die vom Projekt evtl überdeckt werden in den Vordergrund geholt werden 
            ' das muss jedoch nur gemacht werden, wenn nicht vorher schon zeichnePhasenInProjekt oder zeichneMilestonesInProjekt aufgerufen wurde 
            Call bringChartsToFront(projectShape)
        End If



        appInstance.EnableEvents = formerEE
        enableOnUpdate = True

    End Sub


    '' '' ''' <summary>
    '' '' ''' gibt die Anzahl Zeilen zurück, die das Projekt im "expanded View Mode" benötigt 
    '' '' ''' </summary>
    '' '' ''' <param name="hproj"></param>
    '' '' ''' <returns></returns>
    '' '' ''' <remarks></remarks>
    '' ''Public Function calculateNeededLines(ByVal hproj As clsProjekt) As Integer


    '' ''    Dim phasenName As String
    '' ''    Dim zeilenOffset As Integer = 1
    '' ''    Dim lastEndDate As Date = StartofCalendar.AddDays(-1)
    '' ''    Dim tmpValue As Integer


    '' ''    If awinSettings.drawphases Then

    '' ''        For i = 1 To hproj.CountPhases

    '' ''            With hproj.getPhase(i)

    '' ''                phasenName = .name
    '' ''                If DateDiff(DateInterval.Day, lastEndDate, .getStartDate) < 0 Then
    '' ''                    zeilenOffset = zeilenOffset + 1
    '' ''                    lastEndDate = StartofCalendar.AddDays(-1)
    '' ''                End If


    '' ''                If DateDiff(DateInterval.Day, lastEndDate, .getEndDate) > 0 Then
    '' ''                    lastEndDate = .getEndDate
    '' ''                End If

    '' ''            End With


    '' ''        Next

    '' ''        If hproj.CountPhases > 1 Then
    '' ''            tmpValue = zeilenOffset
    '' ''        Else
    '' ''            tmpValue = 1
    '' ''        End If


    '' ''    Else
    '' ''        tmpValue = 1
    '' ''    End If


    '' ''    calculateNeededLines = tmpValue


    '' ''End Function


    ''' <summary>
    ''' gibt die Anzahl Zeilen zurück, die das angegebene Shape benötigt  
    ''' </summary>
    ''' <param name="shpElement"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function getNeededSpace(ByVal shpElement As Excel.Shape) As Integer

        Dim tmpValue As Integer
        Dim zeileTop As Integer, zeileBottom As Integer

        zeileTop = calcYCoordToZeile(shpElement.Top)
        zeileBottom = calcYCoordToZeile(shpElement.Top + shpElement.Height)

        tmpValue = System.Math.Max(zeileBottom - zeileTop, 1)

        getNeededSpace = tmpValue


    End Function


    ''' <summary>
    ''' verschiebt ab Zeile "von" um "anzahlZeilen" alle Projekt-Shapes nach unten 
    ''' aktualisiert auch die tfzeile in den Projekten entsprechend
    ''' die Projekte, die in der collection selCollection enthalten sind, werden nicht nach unten verschoben 
    ''' Projekte, die ab Zeile Stoppzeile stehen, werden nicht nach unten verschoben 
    ''' </summary>
    ''' <param name="anzahlZeilen"></param>
    ''' <remarks></remarks>
    Public Sub moveShapesDown(ByVal selCollection As Collection, _
                              ByVal vonZeile As Integer, ByVal anzahlZeilen As Integer, ByVal stoppzeile As Integer)

        Dim worksheetShapes As Excel.Shapes
        Dim shpElement As Excel.Shape
        Dim formerEOU As Boolean = enableOnUpdate
        Dim shapeType As Integer
        Dim obererRand As Double = calcZeileToYCoord(vonZeile)
        Dim stoppRand As Double = calcZeileToYCoord(stoppzeile)
        Dim differenz As Double = anzahlZeilen * boxHeight
        Dim hproj As clsProjekt

        enableOnUpdate = False

        ' eine maximale Größe, um sicherzugehen, daß in diesem Fall alle Projekte verschoben werden 
        If stoppzeile = 0 Then
            stoppRand = calcZeileToYCoord(40000)
        End If

        Try

            worksheetShapes = CType(appInstance.Worksheets(arrWsNames(3)), Excel.Worksheet).Shapes

        Catch ex As Exception
            Throw New Exception("in moveShapesDown : keine Shapes Zuordnung möglich ")
        End Try


        ' jetzt werden die Shapes verschoben ...
        For i = 1 To worksheetShapes.Count
            shpElement = worksheetShapes.Item(i)

            With shpElement

                If Not CBool(.HasChart) Then

                    shapeType = CInt(.AlternativeText)

                    If .Top >= obererRand And (Not selCollection.Contains(shpElement.Name)) _
                        And .Top < stoppRand Then
                        .Top = CSng(.Top + differenz)

                        ' Ergänzung 11.5.2014: Projekte Anpassen und projectboardShapes Einträge korrigieren 
                        If shapeType = PTshty.phaseE Or shapeType = PTshty.phaseN Or shapeType = PTshty.phase1 Or _
                        shapeType = PTshty.milestoneE Or shapeType = PTshty.milestoneN Or _
                        shapeType = PTshty.status Then

                            projectboardShapes.add(shpElement)

                        ElseIf isProjectType(shapeType) Then

                            projectboardShapes.add(shpElement)
                            hproj = ShowProjekte.getProject(shpElement.Name)
                            'hproj.tfZeile = calcYCoordToZeile(shpElement.Top)
                            hproj.tfZeile = hproj.tfZeile + anzahlZeilen

                        End If

                    End If

                End If

            End With

        Next


        enableOnUpdate = formerEOU

    End Sub

    ''' <summary>
    ''' verschiebt alle Shapes, die ab vonZeile liegen um eins nach oben
    ''' vorher wird aber geprüft, on das frei ist   
    ''' </summary>
    ''' <param name="vonZeile"></param>
    ''' <remarks></remarks>
    Public Sub moveShapesUp(ByVal vonZeile As Integer, ByVal anzahlZeilen As Integer)

        Dim worksheetShapes As Excel.Shapes
        Dim shpElement As Excel.Shape
        Dim formerEOU As Boolean = enableOnUpdate
        Dim zielZeileIsFrei As Boolean = False
        Dim zielZeile As Integer = vonZeile
        Dim shapeType As Integer
        Dim hproj As clsProjekt

        Dim obererRand As Double = calcZeileToYCoord(vonZeile)
        Dim differenz As Double = anzahlZeilen * boxHeight

        enableOnUpdate = False


        If magicBoardZeileIstFrei(zielZeile) Then

            Try

                worksheetShapes = CType(appInstance.Worksheets(arrWsNames(3)), Excel.Worksheet).Shapes

            Catch ex As Exception
                Throw New Exception("in moveShapesUp : keine Shapes Zuordnung möglich ")
            End Try


            ' jetzt werden die Shapes verschoben ...
            For i = 1 To worksheetShapes.Count
                shpElement = worksheetShapes.Item(i)

                With shpElement

                    If Not CBool(.HasChart) Then

                        If .Top >= obererRand Then
                            .Top = CSng(.Top - differenz)

                            ' Ergänzung 11.5.2014: Projekte Anpassen und projectboardShapes Einträge korrigieren 
                            If shapeType = PTshty.phaseE Or shapeType = PTshty.phaseN Or shapeType = PTshty.phase1 Or _
                            shapeType = PTshty.milestoneE Or shapeType = PTshty.milestoneN Or _
                            shapeType = PTshty.status Then

                                projectboardShapes.add(shpElement)

                            ElseIf isProjectType(shapeType) Then

                                projectboardShapes.add(shpElement)
                                hproj = ShowProjekte.getProject(shpElement.Name)
                                hproj.tfZeile = calcYCoordToZeile(shpElement.Top)

                            End If

                        End If

                    End If

                End With

            Next

            ' jetzt werden die Projekte angepasst 
            ' das unten folgende ist durch die Ergänzung oben vom 11.5.14 nicht mehr notwendig 

            'For Each kvp As KeyValuePair(Of String, clsProjekt) In ShowProjekte.Liste

            '    With kvp.Value
            '        If .tfZeile >= vonZeile And .tfZeile - anzahlZeilen >= 2 Then
            '            .tfZeile = .tfZeile - anzahlZeilen
            '        End If
            '    End With

            'Next

        End If



        enableOnUpdate = formerEOU

    End Sub

    ''' <summary>
    ''' zeichnet für das Projekt das Status Shape; wenn es bereits existiert, wird das alte gelöscht, das neue gezeichnet 
    ''' wenn number > 0 , wird diese Zahl in das Symbol geschrieben 
    ''' </summary>
    ''' <param name="hproj"></param>
    ''' <param name="number"></param>
    ''' <remarks></remarks>
    Public Sub zeichneStatusSymbolInPlantafel(ByVal hproj As clsProjekt, ByVal number As Integer)
        Dim top As Double, left As Double, height As Double, width As Double
        Dim worksheetShapes As Excel.Shapes
        Dim statusShape As Excel.Shape
        Dim shpName As String
        Dim timeAtStatus As Date = hproj.timeStamp
        Dim heuteColumn As Integer = getColumnOfDate(timeAtStatus)

        With CType(appInstance.Worksheets(arrWsNames(3)), Excel.Worksheet)
            worksheetShapes = .Shapes

            shpName = projectboardShapes.calcStatusShapeName(hproj.name, heuteColumn)
            ' existiert das schon ? 
            Try
                statusShape = worksheetShapes.Item(shpName)
                statusShape.Delete()
            Catch ex As Exception
                'statusShape = Nothing
            End Try

            'If statusShape Is Nothing Then

            hproj.calculateStatusCoord(timeAtStatus, top, left, width, height)
            statusShape = .Shapes.AddShape(Type:=Microsoft.Office.Core.MsoAutoShapeType.msoShapeOval, _
                                            Left:=CSng(left), Top:=CSng(top), Width:=CSng(width), Height:=CSng(height))

            With statusShape
                .Name = shpName
                .Title = "Status"
                .AlternativeText = CInt(PTshty.status).ToString
            End With

            Call defineStatusAppearance(hproj, number, statusShape)

            ' jetzt der Liste der ProjectboardShapes hinzufügen
            projectboardShapes.add(statusShape)

            'shapesCollection.Add(resultShape.Name)

            'End If


        End With

        ' jetzt müssen die Charts ggf wieder nach vorne gebracht werden 
        Call bringChartsToFront(statusShape)


    End Sub

    ''' <summary>
    ''' wird von der WPFPIE Vorage aufgerufen !
    ''' </summary>
    ''' <param name="nameList"></param>
    ''' <param name="farbTyp"></param>
    ''' <param name="numberIt"></param>
    ''' <remarks></remarks>
    Public Sub zeichneMilestones(ByVal nameList As Collection, ByVal farbTyp As Integer, ByVal numberIt As Boolean)
        ' tue es für alle Projekte in Showprojekte 


        Dim todoListe As New SortedList(Of Long, clsProjekt)
        Dim key As Long
        Dim formerEE As Boolean = appInstance.EnableEvents
        Dim formereO As Boolean = enableOnUpdate

        appInstance.EnableEvents = False
        enableOnUpdate = False

        If selectedProjekte.Count > 0 Then
            For Each kvp As KeyValuePair(Of String, clsProjekt) In selectedProjekte.Liste

                key = 10000 * kvp.Value.tfZeile + kvp.Value.tfspalte
                todoListe.Add(key, kvp.Value)

            Next
        Else
            For Each kvp As KeyValuePair(Of String, clsProjekt) In ShowProjekte.Liste

                key = 10000 * kvp.Value.tfZeile + kvp.Value.tfspalte
                todoListe.Add(key, kvp.Value)

            Next
        End If


        Dim msNumber As Integer = 1

        For Each kvp As KeyValuePair(Of Long, clsProjekt) In todoListe

            Call zeichneMilestonesInProjekt(kvp.Value, nameList, farbTyp, showRangeLeft, showRangeRight, numberIt, msNumber, False)

        Next

        Call awinSelect()

        appInstance.EnableEvents = formerEE
        enableOnUpdate = formereO

    End Sub


    ''' <summary>
    ''' zeichnet die Meilensteine eines Projektes
    ''' </summary>
    ''' <param name="hproj">
    ''' das Projekt, das die Meilensteine enthält</param>
    ''' <param name="namenListe">
    ''' enthält die Namen,der Meilensteine die gezeichnet werden sollen
    ''' wenn leer, werden alle gezeichnet</param>
    ''' <param name="farbTyp">
    ''' gibt an , welche Farbe gezeichnet werden soll; bei 4 werden alle gezeichnet </param>
    ''' <param name="tmpShowrangeleft">
    ''' gibt den linken Rand des Zeitraums an, sofern einer betrachtet werden soll </param>
    ''' <param name="tmpShowrangeRight">gibt den rechten Rand des Zeitraums an, sofern einer betrachtet werden soll </param>
    ''' <param name="numberIt">
    ''' gibt an, ob der Meilenstein nummeriert werden soll</param>
    ''' <param name="msNumber">
    ''' gibt die Nummer an, aber nummeriert werden soll</param>
    ''' <param name="report">
    ''' gibt an, ob vom Reporting aufgerufen
    ''' </param>
    ''' <remarks></remarks>
    Public Sub zeichneMilestonesInProjekt(ByVal hproj As clsProjekt, ByVal namenListe As Collection, ByVal farbTyp As Integer, ByVal tmpShowRangeLeft As Integer, ByVal tmpShowrangeRight As Integer, _
                                                      ByVal numberIt As Boolean, ByRef msNumber As Integer, ByVal report As Boolean)

        Dim top As Double, left As Double, width As Double, height As Double
        Dim resultShape As Excel.Shape
        Dim worksheetShapes As Excel.Shapes
        Dim heute As Date = Date.Now
        Dim alreadyGroup As Boolean = False
        Dim shpElement As Excel.Shape
        Dim shpName As String
        Dim resultColumn As Integer
        Dim onlyFew As Boolean
        Dim projectShape As Excel.Shape
        Dim shapeGruppe As Excel.ShapeRange
        Dim newShape As Excel.ShapeRange = Nothing
        Dim listOFShapes As New Collection
        Dim found As Boolean = True
        Dim showOnlyWithinTimeFrame As Boolean
        Dim vorlagenShape As xlNS.Shape
        Dim realNameList As New Collection

        Try
            If namenListe.Count > 0 Then
                onlyFew = True
                realNameList = hproj.getElemIdsOf(namenListe, True)
            Else
                onlyFew = False
                realNameList = hproj.getAllElemIDs(True)
            End If
        Catch ex As Exception
            onlyFew = False
        End Try

        ' jetzt wurde aus der Liste von Namen / oder IDs gesichert eine Liste von IDs gemacht 


        ' es muss abgefangen werden, daß nicht alle Meilensteine gezeichnet werden, wenn namenListe.count > 0, aber später 
        ' die neue namenliste.count  = 0 ; dass wird dadurch sichergestellt, dass onlyFew bereits vor Bearbeitung / Ersetzung der Namenliste gesetzt ist  


        If tmpShowRangeLeft <= 0 Or _
            tmpShowrangeRight <= 0 Or _
            tmpShowRangeLeft > tmpShowrangeRight Then

            showOnlyWithinTimeFrame = False

        Else

            showOnlyWithinTimeFrame = True

        End If



        With CType(appInstance.Worksheets(arrWsNames(3)), Excel.Worksheet)

            worksheetShapes = .Shapes
            ' Änderung 12.7.14 Alle Milestone Shapes in ein gruppiertes Shape
            ' jetzt muss das Projekt-Shape gesucht werden
            Try
                projectShape = worksheetShapes.Item(hproj.name)
            Catch ex As Exception
                found = False
                projectShape = Nothing
            End Try


            ' found=true bedeutet, dass das Shape bereits angezeigt wird  
            If found Then

                ' jetzt muss die Liste an Shapes aufgebaut werden 
                If projectShape.AlternativeText = CInt(PTshty.projektL).ToString Or _
                    projectShape.AutoShapeType = MsoAutoShapeType.msoShapeRoundedRectangle Then

                    listOFShapes.Add(projectShape.Name)

                Else
                    shapeGruppe = projectShape.Ungroup
                    Dim anzElements As Integer = shapeGruppe.Count

                    Dim i As Integer
                    For i = 1 To anzElements
                        listOFShapes.Add(shapeGruppe.Item(i).Name)
                    Next


                End If

                ' hier muss jetzt ausgenutzt werden, dass man bereits die direkten IDs der Meilensteine hat ... 

                ' es muss aber auch berücksichtigt werden, wenn alle gezeigt werden sollen ... 

                For m As Integer = 1 To realNameList.Count

                    Dim cMilestone As clsMeilenstein = hproj.getMilestoneByID(CStr(realNameList.Item(m)))

                    If Not IsNothing(cMilestone) Then
                        Dim cBewertung As clsBewertung

                        cBewertung = cMilestone.getBewertung(1)
                        resultColumn = getColumnOfDate(cMilestone.getDate)

                        If farbTyp = 4 Or farbTyp = cBewertung.colorIndex Then
                            ' es muss nur etwas gemacht werden , wenn entweder alle Farben gezeichnet werden oder eben die übergebene

                            If (showOnlyWithinTimeFrame And (resultColumn < tmpShowRangeLeft Or resultColumn > tmpShowrangeRight)) Then
                                ' nichts machen 
                            Else
                                Dim zeilenoffset As Integer = 0
                                ' hier die übergeordnete Phase holen ...
                                vorlagenShape = MilestoneDefinitions.getShape(cMilestone.name)
                                Dim factorB2H As Double = vorlagenShape.Width / vorlagenShape.Height

                                hproj.calculateMilestoneCoord(cMilestone.getDate, zeilenoffset, factorB2H, top, left, width, height)
                                'hproj.calculateResultCoord(cResult.getDate, zeilenoffset, top, left, width, height)

                                shpName = projectboardShapes.calcMilestoneShapeName(hproj.name, cMilestone.nameID)

                                ' existiert das schon ? 
                                Try
                                    shpElement = worksheetShapes.Item(shpName)
                                Catch ex As Exception
                                    shpElement = Nothing
                                End Try

                                If shpElement Is Nothing Then

                                    If report Then
                                        top = top - boxWidth
                                    End If

                                    ' Alt - Start 
                                    resultShape = .Shapes.AddShape(Type:=vorlagenShape.AutoShapeType, _
                                                                    Left:=CSng(left), Top:=CSng(top), Width:=CSng(width), Height:=CSng(height))
                                    vorlagenShape.PickUp()
                                    resultShape.Apply()

                                    resultShape.Rotation = vorlagenShape.Rotation

                                    With resultShape
                                        .Name = shpName
                                        .Title = cMilestone.nameID
                                        .AlternativeText = CInt(PTshty.milestoneN).ToString
                                    End With
                                    ' Alt - Ende



                                    msNumber = msNumber + 1
                                    If numberIt Then
                                        Call defineResultAppearance(hproj, msNumber, resultShape, cBewertung)

                                    Else
                                        Call defineResultAppearance(hproj, 0, resultShape, cBewertung)
                                    End If

                                    ' jetzt der Liste der ProjectboardShapes hinzufügen
                                    projectboardShapes.add(resultShape)

                                    ' jetzt der Liste von Shapes hinzufügen, die dann nachher zum ProjektShape gruppiert werden sollen 
                                    listOFShapes.Add(resultShape.Name)

                                End If

                            End If
                        End If
                    End If



                Next


            End If


            If listOFShapes.Count > 1 Then
                ' hier werden die Shapes gruppiert
                projectShape = projectboardShapes.groupShapes(listOFShapes, hproj.name)

                ' jetzt der Liste der ProjectboardShapes hinzufügen
                projectboardShapes.add(projectShape)
            End If


        End With


        ' jetzt müssen ggf die Charts wieder in den Vordergrund gebracht werden 
        Call bringChartsToFront(projectShape)

    End Sub

    ''' <summary>
    ''' zeichnet die Werte der Rollen und Kosten auf die Projekt-Tafel
    ''' </summary>
    ''' <param name="hproj">das Projekt, das gezeichnet werden soll </param>
    ''' <param name="namenListe">die Liste der Rollen bzw. Kosten</param>
    ''' <param name="tmpShowRangeLeft">linke Spalte des Bereiches, in dem gezeichnet werden soll</param>
    ''' <param name="tmpShowrangeRight">rechte Spalte des Bereiches, in dem gezeichnet werden soll</param>
    ''' <param name="type">gibt an den Type an, damit lässt sich entscheiden, ob Rolle / Kosten in der Namenliste stehen</param>
    ''' <remarks></remarks>
    Public Sub zeichneRollenKostenWerteInProjekt(ByVal hproj As clsProjekt, ByVal namenListe As Collection, ByVal tmpShowRangeLeft As Integer, ByVal tmpShowrangeRight As Integer, _
                                                          ByVal type As String)

        ' aktuell wird das nur im Fall nicht-extended Mode angezeigt 
        If awinSettings.drawphases Then
            ' nichts tun 
            Call MsgBox("wird aktuell nur im Einzeilen - Modus unterstützt" & vbLf & _
                         "Wählen Sie Extended Mode = Nein")
            Exit Sub
        End If

        ' aktuell wird nur unterstützt, einen Monat anzuzeigen 
        If tmpShowRangeLeft <> tmpShowrangeRight Then
            Call MsgBox("aktuell wird nur ein Monat unterstützt")
            Exit Sub
        End If

        ' bestimme die Zeile und die Spalte 
        Dim currentRow As Integer = hproj.tfZeile + 1
        Dim currentColumn As Integer = tmpShowRangeLeft

        ' bestimme den Wert
        Dim currentValue As Double = hproj.getBedarfeInMonth(namenListe, type, tmpShowRangeLeft)

        ' schreibe jetzt den Wert in die Zelle
        Dim formerEE As Boolean = appInstance.EnableEvents
        appInstance.EnableEvents = False

        With CType(appInstance.Worksheets(arrWsNames(3)), Excel.Worksheet)
            If currentValue > 0 Then
                .Cells(currentRow, currentColumn).value = CInt(currentValue)

            End If

        End With

        appInstance.EnableEvents = formerEE


    End Sub

    ''' <summary>
    ''' trägt bei Projektlinie den Namen ein ... 
    ''' </summary>
    ''' <param name="hproj"></param>
    ''' <remarks></remarks>
    Public Sub zeichneNameInProjekt(ByVal hproj As clsProjekt)

        Dim projectTop As Single, projectLeft As Single, projectHeight As Single, projectWidth As Single
        Dim txtTop As Single, txtLeft As Single, txtwidth As Single, txtHeight As Single
        Dim pNameShape As Excel.Shape
        Dim worksheetShapes As Excel.Shapes

        Dim projectShape As Excel.Shape
        Dim shapeGruppe As Excel.ShapeRange

        Dim listOFShapes As New Collection
        Dim found As Boolean = True


        With CType(appInstance.Worksheets(arrWsNames(3)), Excel.Worksheet)

            worksheetShapes = .Shapes
            ' Änderung 12.7.14 Alle Milestone Shapes in ein gruppiertes Shape
            ' jetzt muss das Projekt-Shape gesucht werden
            Try
                projectShape = worksheetShapes.Item(hproj.name)
            Catch ex As Exception
                found = False
                projectShape = Nothing
            End Try


            ' found=true bedeutet, dass das Shape bereits angezeigt wird  
            If found Then



                ' Merken der Koordinaten 
                ' bestimmen der Text Koordinaten 
                With projectShape
                    projectTop = .Top
                    projectLeft = .Left
                    projectWidth = .Width
                    projectHeight = .Height
                End With

                txtTop = projectTop
                txtLeft = projectLeft + 7
                txtwidth = 30
                txtHeight = 30

                ' jetzt muss die Liste an Shapes aufgebaut werden 

                If projectShape.AlternativeText = CInt(PTshty.projektL).ToString Or _
                    projectShape.AutoShapeType = MsoAutoShapeType.msoShapeRoundedRectangle Then
                    listOFShapes.Add(projectShape.Name)
                Else
                    shapeGruppe = projectShape.Ungroup
                    Dim anzElements As Integer = shapeGruppe.Count
                    ' hier muss der alte Shape Text rausgelöscht werdewn 

                    Dim oldTxtxShape As Excel.Shape = Nothing
                    For Each tmpshape As Excel.Shape In shapeGruppe
                        If tmpshape.AlternativeText = "(Projektname)" Then
                            oldTxtxShape = tmpshape
                        Else
                            listOFShapes.Add(tmpshape.Name)
                        End If
                    Next

                    ' jetzt muss der alte Text gelöscht werden ...
                    If Not IsNothing(oldTxtxShape) Then
                        oldTxtxShape.Delete()
                    End If


                End If



                ' ab jetzt darf auf projectShape nicht mehr zugegriffen werden, da es ggf bereits im Else-Zweig aufgelöst wurde ...


                ' jetzt muss das Textshape erzeugt werden 
                pNameShape = worksheetShapes.AddLabel(MsoTextOrientation.msoTextOrientationHorizontal, _
                                                        txtLeft, txtTop, txtwidth, txtHeight)

                With pNameShape
                    .AlternativeText = "(Projektname)"
                    .TextFrame2.AutoSize = MsoAutoSize.msoAutoSizeShapeToFitText
                    .TextFrame2.WordWrap = MsoTriState.msoFalse
                    .TextFrame2.TextRange.Text = hproj.getShapeText
                    .TextFrame2.TextRange.Font.Size = hproj.Schrift
                    .TextFrame2.MarginLeft = 0
                    .TextFrame2.MarginRight = 0
                    .TextFrame2.MarginTop = 0
                    .TextFrame2.MarginBottom = 0
                    .TextFrame2.VerticalAnchor = MsoVerticalAnchor.msoAnchorMiddle
                    .TextFrame2.HorizontalAnchor = MsoHorizontalAnchor.msoAnchorCenter

                    .Fill.Visible = MsoTriState.msoTrue
                    .Fill.ForeColor.RGB = RGB(255, 255, 255)
                    .Fill.Transparency = 0
                    .Fill.Solid()

                End With

                ' jetzt muss das Shape noch in der Höhe richtig positioniert werden 
                Dim diff As Single
                If awinSettings.drawphases Then
                    diff = CSng(0.3 * boxHeight)
                Else
                    diff = (pNameShape.Height - projectHeight) / 2
                End If
                pNameShape.Top = projectTop - diff

                If pNameShape.Width > projectWidth Then
                    pNameShape.TextFrame2.TextRange.Text = ""
                End If

                ' jetzt wird das Shape aufgenommen 
                listOFShapes.Add(pNameShape.Name)


            End If


            If listOFShapes.Count > 1 Then
                ' hier werden die Shapes gruppiert
                projectShape = projectboardShapes.groupShapes(listOFShapes, hproj.name)

                ' jetzt der Liste der ProjectboardShapes hinzufügen
                projectboardShapes.add(projectShape)
            End If


        End With


        ' jetzt müssen ggf die Charts wieder in den Vordergrund gebracht werden 
        Call bringChartsToFront(projectShape)


    End Sub

    ''' <summary>
    ''' zeichnet die Abhängigkeiten zu dem übergebenen Projekt 
    ''' </summary>
    ''' <param name="hproj">Projekt, dessen Abhängigkeiten dargestellt werden sollen</param>
    ''' <param name="type">welche Art Abhängigkeit soll dargestellt werden</param>
    ''' <param name="auswahl">0: sowohl incoming als outgoing Abhängigkeiten
    ''' 1: nur outgoing Abhängigkeiten
    ''' 2: nur incoming abhängigkeiten</param>
    ''' <remarks></remarks>
    Public Sub zeichneDependenciesOfProject(ByVal hproj As clsProjekt, ByVal type As Integer, ByVal auswahl As Integer)

        Dim listeDep As Collection ' nimmt die Liste der abhängigen Projekte auf
        Dim depListe As Collection ' nimmt die Liste der Projekte auf, von denen hproj abhängig ist 
        Dim pShape As Excel.Shape
        Dim dpShape As Excel.Shape
        Dim newConnector As Excel.Shape
        Dim X1, X2, Y1, Y2 As Single
        Dim dProj As clsProjekt
        Dim curDependency As clsDependency

        Dim pName As String = hproj.name, dpName As String

        Dim tmpshapes As Excel.Shapes
        Dim formerEE As Boolean = appInstance.EnableEvents
        Dim formerEOU As Boolean = enableOnUpdate



        listeDep = allDependencies.activeListe(hproj.name, PTdpndncyType.inhalt)
        depListe = allDependencies.passiveListe(hproj.name, PTdpndncyType.inhalt)

        If listeDep.Count = 0 And depListe.Count = 0 Then
            ' es gibt keine Abhängigkeiten
            Throw New Exception("keine Abhängigkeiten vorhanden")
        Else
            ' jetzt werden die Abhängigkeiten gezeichnet ...


            ' Event Behandlung ausschalten 
            enableOnUpdate = False
            appInstance.EnableEvents = False

            tmpshapes = CType(appInstance.Worksheets(arrWsNames(3)), Excel.Worksheet).Shapes
            pShape = tmpshapes.Item(pName)

            ' outgoing dependencies
            If auswahl = 0 Or auswahl = 1 Then

                For d = 1 To listeDep.Count

                    Try
                        dpName = CStr(listeDep.Item(d))
                        dpShape = tmpshapes.Item(dpName)
                        dProj = ShowProjekte.getProject(dpName)
                        Dim curDegree As Integer
                        curDependency = allDependencies.getDependency(PTdpndncyType.inhalt, pName, dpName)
                        If Not IsNothing(curDependency) Then
                            curDegree = curDependency.degree
                        Else
                            curDegree = PTdpndncy.schwach
                        End If

                        'Dim newShapeName As String = pName.Trim & "#" & dpName.Trim
                        Dim newShapeName As String = projectboardShapes.calcDependencyShapeName(pName, dpName)

                        ' prüfen , ob das Shape schon existiert ? 
                        Try

                            newConnector = tmpshapes.Item(newShapeName)

                            With newConnector
                                If curDegree = PTdpndncy.schwach Then
                                    .Line.Weight = 4.0
                                    .Line.DashStyle = MsoLineDashStyle.msoLineLongDash
                                Else
                                    .Line.Weight = 4.0
                                    .Line.DashStyle = MsoLineDashStyle.msoLineSolid
                                End If
                            End With



                        Catch ex As Exception

                            Call calculateDepCoord(pShape, dpShape, X1, Y1, X2, Y2)
                            newConnector = tmpshapes.AddConnector(MsoConnectorType.msoConnectorStraight, X1, Y1, X2, Y2)

                            With newConnector
                                .Line.EndArrowheadStyle = MsoArrowheadStyle.msoArrowheadTriangle
                                .ConnectorFormat.BeginConnect(pShape, 3)
                                .ConnectorFormat.EndConnect(dpShape, 1)
                                .Line.ForeColor.RGB = CInt(awinSettings.AmpelRot)
                                If curDegree = PTdpndncy.schwach Then
                                    .Line.Weight = 4.0
                                    .Line.DashStyle = MsoLineDashStyle.msoLineLongDash
                                Else
                                    .Line.Weight = 4.0
                                    .Line.DashStyle = MsoLineDashStyle.msoLineSolid
                                End If
                                .Name = newShapeName
                                .AlternativeText = CInt(PTshty.dependency).ToString
                            End With

                            Call bringChartsToFront(newConnector)

                        End Try



                    Catch ex As Exception

                    End Try

                Next


            End If

            ' incoming dependencies
            dpShape = tmpshapes.Item(pName)
            dpName = pName
            dProj = hproj
            If auswahl = 0 Or auswahl = 2 Then

                For d = 1 To depListe.Count

                    Try
                        pName = CStr(depListe.Item(d))
                        pShape = tmpshapes.Item(pName)
                        hproj = ShowProjekte.getProject(pName)

                        Dim curDegree As Integer
                        curDependency = allDependencies.getDependency(PTdpndncyType.inhalt, pName, dpName)
                        If Not IsNothing(curDependency) Then
                            curDegree = curDependency.degree
                        Else
                            curDegree = PTdpndncy.schwach
                        End If

                        Dim newShapeName As String = projectboardShapes.calcDependencyShapeName(pName, dpName)
                        'Dim newShapeName As String = pName.Trim & "#" & dpName.Trim

                        ' prüfen , ob das Shape schon existiert ? 
                        Try

                            newConnector = tmpshapes.Item(newShapeName)

                            With newConnector
                                If curDegree = PTdpndncy.schwach Then
                                    .Line.Weight = 4.0
                                    .Line.DashStyle = MsoLineDashStyle.msoLineLongDash
                                Else
                                    .Line.Weight = 4.0
                                    .Line.DashStyle = MsoLineDashStyle.msoLineSolid
                                End If
                            End With


                        Catch ex As Exception

                            Call calculateDepCoord(pShape, dpShape, X1, Y1, X2, Y2)
                            newConnector = tmpshapes.AddConnector(MsoConnectorType.msoConnectorStraight, X1, Y1, X2, Y2)

                            With newConnector
                                .Line.EndArrowheadStyle = MsoArrowheadStyle.msoArrowheadTriangle
                                .ConnectorFormat.BeginConnect(pShape, 3)
                                .ConnectorFormat.EndConnect(dpShape, 1)
                                .Line.ForeColor.RGB = CInt(awinSettings.AmpelRot)
                                If curDegree = PTdpndncy.schwach Then
                                    .Line.Weight = 4.0
                                    .Line.DashStyle = MsoLineDashStyle.msoLineLongDash
                                Else
                                    .Line.Weight = 4.0
                                    .Line.DashStyle = MsoLineDashStyle.msoLineSolid
                                End If
                                .Name = newShapeName
                                .Title = "Dependency"
                            End With

                            Call bringChartsToFront(newConnector)

                        End Try

                    Catch ex As Exception

                    End Try

                Next


            End If

            ' Event Behandlung auf vorherigen Zustand setzen ...
            appInstance.EnableEvents = formerEE
            enableOnUpdate = formerEOU
        End If
    End Sub


    ''' <summary>
    ''' berechnet die Koordinaten des Abhängigkeit-Konnektors - der Linie
    ''' </summary>
    ''' <param name="pShape"></param>
    ''' <param name="dpShape"></param>
    ''' <param name="X1"></param>
    ''' <param name="Y1"></param>
    ''' <param name="X2"></param>
    ''' <param name="Y2"></param>
    ''' <remarks></remarks>
    Public Sub calculateDepCoord(ByVal pShape As Excel.Shape, ByVal dpShape As Excel.Shape, _
                                     ByRef X1 As Single, ByRef Y1 As Single, ByRef X2 As Single, ByRef Y2 As Single)

        With pShape
            X1 = .Left + .Width / 2
            Y1 = .Top + .Height
        End With

        With dpShape
            X2 = .Left + .Width / 2
            Y2 = .Top
        End With

    End Sub

    ' nicht mehr notwendig, da eine zeichnePhasenInProjekt mit optionalen Paramtern das selbe erledigt 
    ' ''' <summary>
    ' ''' zeichnet für das angegebene Projekt hproj alle in namenliste enthaltenen Phasen
    ' ''' wenn namenliste leer ist, werden alle Phasen des Projekts gezeichnet
    ' ''' numberit steuet, ob die Phase für Reporting Zwecke eine Nummerierung erhalten soll 
    ' ''' </summary>
    ' ''' <param name="hproj"></param>
    ' ''' <param name="namenListe"></param>
    ' ''' <param name="numberIt"></param>
    ' ''' <param name="msNumber"></param>
    ' ''' <remarks></remarks>
    'Public Sub zeichnePhasenInProjekt(ByVal hproj As clsProjekt, ByVal namenListe As Collection, ByVal numberIt As Boolean, ByRef msNumber As Integer)

    '    'Dim top1 As Double, left1 As Double, top2 As Double, left2 As Double
    '    Dim top As Double, left As Double, width As Double, height As Double
    '    Dim nummer As Integer, gesamtZahl As Integer
    '    Dim phasenShape As Excel.Shape
    '    Dim worksheetShapes As Excel.Shapes
    '    Dim heute As Date = Date.Now
    '    Dim alreadyGroup As Boolean = False
    '    Dim shpElement As Excel.Shape
    '    Dim vorlagenShape As xlNS.Shape
    '    Dim shpName As String

    '    Dim onlyFew As Boolean
    '    Dim projectShape As Excel.Shape
    '    Dim shapeGruppe As Excel.ShapeRange
    '    Dim listOFShapes As New Collection
    '    Dim found As Boolean = True


    '    Dim nameIstInListe As Boolean
    '    Dim linienDicke As Double = 2.0


    '    ' als wievielte Phase wird das Shape gezeichnet ... 
    '    nummer = 1

    '    ' sollen nur die in der Namenliste aufgeführten Phasen gezeichnet werden ? 
    '    Try
    '        If namenListe.Count > 0 Then
    '            onlyFew = True
    '            gesamtZahl = namenListe.Count
    '        Else
    '            onlyFew = False
    '            gesamtZahl = hproj.CountPhases
    '        End If
    '    Catch ex As Exception
    '        onlyFew = False
    '    End Try


    '    With CType(appInstance.Worksheets(arrWsNames(3)), Excel.Worksheet)

    '        worksheetShapes = .Shapes

    '        ' Änderung 12.7.14 Alle Phasen Shapes in ein gruppiertes Shape
    '        ' jetzt muss das Projekt-Shape gesucht werden
    '        Try
    '            projectShape = worksheetShapes.Item(hproj.name)
    '        Catch ex As Exception
    '            found = False
    '            projectShape = Nothing
    '        End Try

    '        ' nur wenn das Projektshape überhaupt existiert, wird gezeichnet 
    '        If found Then


    '            ' jetzt muss die Liste an Shapes aufgebaut werden 
    '            If projectShape.AutoShapeType = MsoAutoShapeType.msoShapeRoundedRectangle Then

    '                listOFShapes.Add(projectShape.Name)

    '            Else
    '                shapeGruppe = projectShape.Ungroup
    '                Dim anzElements As Integer = shapeGruppe.Count

    '                Dim i As Integer
    '                For i = 1 To anzElements
    '                    listOFShapes.Add(shapeGruppe.Item(i).Name)
    '                Next
    '            End If


    '            For p = 1 To hproj.CountPhases

    '                Dim cphase As clsPhase = hproj.getPhase(p)

    '                Try
    '                    nameIstInListe = namenListe.Contains(cphase.name)
    '                Catch ex As Exception
    '                    nameIstInListe = False
    '                End Try



    '                If onlyFew And Not nameIstInListe Then
    '                    ' nichts machen 
    '                Else

    '                    linienDicke = boxHeight * 0.3
    '                    vorlagenShape = PhaseDefinitions.getShape(cphase.name)

    '                    Try
    '                        'cphase.calculateLineCoord(hproj.tfZeile, nummer, gesamtZahl, top1, left1, top2, left2, linienDicke)
    '                        cphase.calculatePhaseShapeCoord(top, left, width, height)

    '                    Catch ex As Exception
    '                        Throw New ArgumentException(ex.Message)
    '                    End Try

    '                    nummer = nummer + 1

    '                    shpName = projectboardShapes.calcPhaseShapeName(hproj.name, cphase.nameID)

    '                    Try
    '                        shpElement = worksheetShapes.Item(shpName)
    '                    Catch ex As Exception
    '                        shpElement = Nothing
    '                    End Try

    '                    If shpElement Is Nothing Then


    '                        'phasenShape = .Shapes.AddConnector(MsoConnectorType.msoConnectorStraight, CSng(left1), CSng(top1), CSng(left2), CSng(top2))
    '                        phasenShape = .Shapes.AddShape(Type:=vorlagenShape.AutoShapeType, _
    '                                                                Left:=CSng(left), Top:=CSng(top), Width:=CSng(width), Height:=CSng(height))
    '                        vorlagenShape.PickUp()
    '                        phasenShape.Apply()

    '                        With phasenShape
    '                            .Name = shpName
    '                            .Title = cphase.nameID
    '                            .AlternativeText = CInt(PTshty.phaseN).ToString
    '                        End With

    '                        msNumber = msNumber + 1
    '                        'If numberIt Then
    '                        '    Call defineLineAppearance(hproj, cphase, msNumber, phasenShape, linienDicke)

    '                        'Else
    '                        '    Call defineLineAppearance(hproj, cphase, 0, phasenShape, linienDicke)
    '                        'End If

    '                        ' jetzt der Liste der ProjectboardShapes hinzufügen
    '                        projectboardShapes.add(phasenShape)

    '                        ' jetzt der Liste von Shapes hinzufügen, die dann nachher zum ProjektShape gruppiert werden sollen 
    '                        listOFShapes.Add(phasenShape.Name)

    '                    End If


    '                End If

    '            Next

    '        End If



    '        If listOFShapes.Count > 1 Then
    '            ' hier werden die Shapes gruppiert
    '            projectShape = projectboardShapes.groupShapes(listOFShapes, hproj.name)

    '            ' jetzt der Liste der ProjectboardShapes hinzufügen
    '            projectboardShapes.add(projectShape)

    '        End If

    '    End With


    '    ' jetzt müssen die Charts ggf wieder nach vorne gebracht werden 
    '    Call bringChartsToFront(projectShape)

    'End Sub

    ''' <summary>
    ''' aktualisiert mit dem angegebenen Projekt die evtl angezeigten Info Forms zu Phase, Meilenstein oder Status
    ''' wenn das Objekt mit dem Namen nicht existiert, dann wird es entsprechend dort vermerkt 
    ''' </summary>
    ''' <param name="hproj"></param>
    ''' <remarks></remarks>
    Public Sub aktualisierePMSForms(ByVal hproj As clsProjekt)

        Dim phaseNameID As String
        Dim milestoneNameID As String

        If formPhase.Visible Then
            phaseNameID = formPhase.phaseNameID
            Call updatePhaseInformation(hproj, phaseNameID)
        End If

        If formMilestone.Visible Then

            milestoneNameID = formMilestone.milestoneNameID
            Call updateMilestoneInformation(hproj, milestoneNameID)

        End If



        If formStatus.Visible Then

            Call zeichneStatusSymbolInPlantafel(hproj, 0)
            Call updateStatusInformation(hproj)

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


        If Not (hproj Is Nothing) Then

            With appInstance.Worksheets(arrWsNames(3))
                Dim tmpArray() As String
                Dim anzDiagrams As Integer
                anzDiagrams = CType(.Chartobjects, Excel.ChartObjects).Count

                If anzDiagrams > 0 Then
                    For i = 1 To anzDiagrams
                        chtobj = CType(.ChartObjects(i), Excel.ChartObject)
                        If chtobj.Name <> "" Then
                            tmpArray = chtobj.Name.Split(New Char() {CType("#", Char)}, 5)
                            ' chtobj name ist aufgebaut: pr#PTprdk.kennung#pName#Auswahl
                            If tmpArray(0) = "pr" Then

                                ' tk/ur: 2.7.15 das muss nochmal in Ruhe überarbeitet werden 
                                ' Aufnahme Diagramme 
                                'ur:12.03.2015
                                ' Diagramlist auf den neuesten Stand bringen, damit der Resize der Charts funktioniert

                                ' '' ''founddiagram = DiagramList.getDiagramm(chtobj.Name)
                                ' '' ''DiagramList.Remove(chtobj.Name)
                                ' '' ''With founddiagram
                                ' '' ''    tmpArray(2) = vglName
                                ' '' ''    IDkennung = Join(tmpArray, "#")
                                ' '' ''    .kennung = IDkennung
                                ' '' ''End With
                                ' '' ''DiagramList.Add(founddiagram)
                                ' VORSICHT: das Diagram 'founddiagram' ist von den Inhalten in der DiagramList inkonsistenz.
                                '           DiagramTitle und die myCollection stimmen nicht mit dem selektierten Projekt überein.
                                ' TODO: den in den update-Routinen zusammengesetzen DiagramTitle und die aktuelle myCollection müssen noch in das ListenElement richtig eingetragen werden.
                                ' siehe JIRA PT89
                                ' ur:12.03.2025: ende

                                If replaceProj Or (tmpArray(2).Trim = vglName) Then
                                    Select Case tmpArray(1)


                                        ' replaceProj sorgt in den nachfolgenden Sequenzen dafür, daß das Chart im Falle eines Aufrufes aus der 
                                        ' Time-Machine (replaceProj = false) nicht in der Skalierung angepasst wird; das geschieht initial beim Laden der Time-Machine
                                        ' wenn es aus dem Selektieren von Projekten aus aufgerufen wird, dann wird die optimal passende Skalierung schon jedesmal berechnet 

                                        Case CInt(PTprdk.Phasen).ToString
                                            ' Update Phasen Diagramm

                                            If CInt(tmpArray(3)) = PThis.current Then
                                                ' nur dann muss aktualisiert werden ...
                                                Call updatePhasesBalken(hproj, chtobj, CInt(tmpArray(3)), replaceProj)
                                            End If


                                        Case CInt(PTprdk.PersonalBalken).ToString

                                            Call updateRessBalkenOfProject(hproj, chtobj, CInt(tmpArray(3)), replaceProj)


                                        Case CInt(PTprdk.PersonalPie).ToString


                                            ' Update Pie-Diagramm
                                            Call updateRessPieOfProject(hproj, chtobj, CInt(tmpArray(3)))


                                        Case CInt(PTprdk.KostenBalken).ToString


                                            Call updateCostBalkenOfProject(hproj, chtobj, CInt(tmpArray(3)), replaceProj)


                                        Case CInt(PTprdk.KostenPie).ToString


                                            Call updateCostPieOfProject(hproj, chtobj, CInt(tmpArray(3)))


                                        Case CInt(PTprdk.StrategieRisiko).ToString

                                            Call updateProjectPfDiagram(hproj, chtobj, CInt(tmpArray(3)))

                                        Case CInt(PTprdk.FitRisikoVol).ToString

                                            Call updateProjectPfDiagram(hproj, chtobj, CInt(tmpArray(3)))

                                        Case CInt(PTprdk.ComplexRisiko).ToString

                                            Call updateProjectPfDiagram(hproj, chtobj, CInt(tmpArray(3)))

                                        Case CInt(PTprdk.Ergebnis).ToString
                                            ' Update Ergebnis Diagramm
                                            Call updateProjektErgebnisCharakteristik2(hproj, chtobj, CInt(tmpArray(3)), replaceProj)

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



    ''' <summary>
    ''' zeichnet für das angegebene Projekt hproj alle in namenliste enthaltenen Phasen, sofern die Phase innerhalb 
    ''' der vonMonth, bisMonth aufgespannten Grenzen liegt 
    ''' wenn namenliste leer ist, werden alle Phasen des Projekts gezeichnet
    ''' numberit steuet, ob die Phase für Reporting Zwecke eine Nummerierung erhalten soll 
    ''' </summary>
    ''' <param name="hproj">das Projekt-Objekt</param>
    ''' <param name="namenListe">Liste der Phasen, die gezeichnet werden sollen</param>
    ''' <param name="vonMonth">linker rand des Kalenderzeitraums, der betrachtet werden soll</param>
    ''' <param name="bisMonth">rechter Rand des Kalenderzeitraums, der betrachtet werden soll</param>
    ''' <param name="numberIt">soll nummeriert werden </param>
    ''' <param name="msNumber">Start der Nummerierung</param>
    ''' <remarks></remarks>
    Public Sub zeichnePhasenInProjekt(ByVal hproj As clsProjekt, ByVal namenListe As Collection, _
                                      ByVal numberIt As Boolean, ByRef msNumber As Integer, _
                                      Optional ByVal vonMonth As Integer = 0, Optional ByVal bisMonth As Integer = 0)

        'Dim top1 As Double, left1 As Double, top2 As Double, left2 As Double
        Dim top As Double, left As Double, width As Double, height As Double
        Dim nummer As Integer
        Dim phasenShape As xlNS.Shape
        Dim worksheetShapes As xlNS.Shapes
        Dim heute As Date = Date.Now
        Dim alreadyGroup As Boolean = False
        Dim shpElement As xlNS.Shape
        Dim vorlagenshape As xlNS.Shape
        Dim shpName As String
        Dim todoListe As New Collection
        Dim realNameList As New Collection
        Dim phasenSchriftgroesse As Double = 5.0

        Dim onlyFew As Boolean
        Dim projectShape As xlNS.Shape
        Dim shapeGruppe As xlNS.ShapeRange
        Dim listOFShapes As New Collection
        Dim found As Boolean = True


        Dim linienDicke As Double = 2.0
        Dim ok As Boolean = True

        ' alle Phasen auslesen , die NameIDs dazu holen 
        Try
            If namenListe.Count > 0 Then
                onlyFew = True
                realNameList = hproj.getElemIdsOf(namenListe, False)
            Else
                onlyFew = False
                realNameList = hproj.getAllElemIDs(False)
            End If
        Catch ex As Exception
            onlyFew = False
        End Try


        ' als wievielte Phase wird das Shape gezeichnet ... 
        nummer = 1




        Try
            If vonMonth = 0 Or bisMonth = 0 Then
                ' alle Phasen betrachten 
                todoListe = realNameList
            Else
                'bringt eine List von Phasen ElemIDs zurück, die den angegebenen Zeitraum berühren / überdecken
                todoListe = hproj.withinTimeFrame(False, vonMonth, bisMonth, realNameList)
            End If



        Catch ex As Exception

        End Try

        With CType(appInstance.Worksheets(arrWsNames(3)), Excel.Worksheet)

            worksheetShapes = .Shapes

            ' Änderung 12.7.14 Alle Phasen Shapes in ein gruppiertes Shape
            ' jetzt muss das Projekt-Shape gesucht werden
            Try
                projectShape = worksheetShapes.Item(hproj.name)
            Catch ex As Exception
                found = False
                projectShape = Nothing
            End Try


            If found Then

                ' jetzt muss die Liste an Shapes aufgebaut werden 
                If projectShape.AlternativeText = CInt(PTshty.projektL).ToString Or _
                    projectShape.AutoShapeType = MsoAutoShapeType.msoShapeRoundedRectangle Then

                    listOFShapes.Add(projectShape.Name)

                Else
                    shapeGruppe = projectShape.Ungroup
                    Dim anzElements As Integer = shapeGruppe.Count

                    Dim i As Integer
                    For i = 1 To anzElements
                        listOFShapes.Add(shapeGruppe.Item(i).Name)
                    Next
                End If


                Dim cphase As clsPhase

                ' in der todoListe stehen jetzt nur Phasen, die den angegeben Zeitraum betreffen 
                For p = 1 To todoListe.Count

                    Dim phaseNameID As String = CStr(todoListe(p))

                    If realNameList.Contains(phaseNameID) Then

                        cphase = hproj.getPhaseByID(phaseNameID)

                        vorlagenshape = PhaseDefinitions.getShape(elemNameOfElemID(phaseNameID))
                        linienDicke = boxHeight * 0.3

                        Try
                            'cphase.calculateLineCoord(hproj.tfZeile, nummer, gesamtZahl, top1, left1, top2, left2, linienDicke)
                            cphase.calculatePhaseShapeCoord(top, left, width, height)
                        Catch ex As Exception
                            ok = False
                        End Try



                        If ok Then
                            nummer = nummer + 1

                            shpName = projectboardShapes.calcPhaseShapeName(hproj.name, cphase.nameID)
                            'shpName = hproj.name & "#" & cphase.name
                            ' existiert das schon ? 
                            Try
                                shpElement = worksheetShapes.Item(shpName)
                            Catch ex As Exception
                                shpElement = Nothing
                            End Try

                            If shpElement Is Nothing Then


                                'phasenShape = .Shapes.AddConnector(MsoConnectorType.msoConnectorStraight, CSng(left1), CSng(top1), CSng(left2), CSng(top2))

                                phasenShape = .Shapes.AddShape(Type:=vorlagenshape.AutoShapeType, _
                                                                    Left:=CSng(left), Top:=CSng(top), Width:=CSng(width), Height:=CSng(height))
                                vorlagenshape.PickUp()
                                phasenShape.Apply()

                                With phasenShape
                                    .Name = shpName
                                    .Title = cphase.nameID
                                    .AlternativeText = CInt(PTshty.phaseN).ToString
                                End With

                                msNumber = msNumber + 1
                                'If numberIt Then
                                '    Call defineLineAppearance(hproj, cphase, msNumber, phasenShape, linienDicke)

                                'Else
                                '    Call defineLineAppearance(hproj, cphase, 0, phasenShape, linienDicke)
                                'End If

                                ' jetzt der Liste der ProjectboardShapes hinzufügen
                                projectboardShapes.add(phasenShape)

                                ' jetzt der Liste von Shapes hinzufügen, die dann nachher zum ProjektShape gruppiert werden sollen 
                                listOFShapes.Add(phasenShape.Name)


                            End If

                        End If

                    End If

                    ok = True

                Next

            End If

            If listOFShapes.Count > 1 Then
                ' hier werden die Shapes gruppiert
                projectShape = projectboardShapes.groupShapes(listOFShapes, hproj.name)

                ' jetzt der Liste der ProjectboardShapes hinzufügen
                projectboardShapes.add(projectShape)

            End If

        End With

        ' jetzt müssen die Charts ggf wieder nach vorne gebracht werden 
        Call bringChartsToFront(projectShape)


    End Sub

    Public Sub defineStatusAppearance(ByVal myproject As clsProjekt, ByVal number As Integer, ByRef myShape As Excel.Shape)
        Dim pColor As Long

        With myproject

            If .ampelStatus = 0 Then
                pColor = awinSettings.AmpelNichtBewertet
            ElseIf .ampelStatus = 1 Then
                pColor = awinSettings.AmpelGruen
            ElseIf .ampelStatus = 2 Then
                pColor = awinSettings.AmpelGelb
            Else
                pColor = awinSettings.AmpelRot
            End If

        End With

        With myShape

            With .Line
                .Visible = Microsoft.Office.Core.MsoTriState.msoTrue
                'If pMarge < 0 Then
                '    .ForeColor.RGB = RGB(255, 0, 0)
                '    .Weight = 2.0
                'Else
                '    .ForeColor.RGB = pcolor
                'End If
                .ForeColor.RGB = CInt(pColor)
                .Transparency = 0
            End With

            With .Fill
                '.Visible = msoTrue
                .ForeColor.RGB = CInt(pColor)
                .ForeColor.TintAndShade = 0
                .ForeColor.Brightness = -0.25

                If roentgenBlick.isOn Then
                    .Transparency = 0.8
                Else
                    .Transparency = 0.0
                End If

                .Solid()

            End With

            Try

                If .TextFrame2.HasText <> MsoTriState.msoFalse Then
                    .TextFrame2.TextRange.Text = ""
                End If

                If number > 0 And Not roentgenBlick.isOn Then

                    With .TextFrame2
                        .MarginLeft = 0
                        .MarginRight = 0
                        .MarginBottom = 0
                        .MarginTop = 0
                        .WordWrap = Microsoft.Office.Core.MsoTriState.msoFalse
                        .VerticalAnchor = MsoVerticalAnchor.msoAnchorMiddle
                        .HorizontalAnchor = MsoHorizontalAnchor.msoAnchorCenter
                        .TextRange.Text = number.ToString
                        .TextRange.Font.Size = awinSettings.fontsizeLegend
                        .TextRange.Font.Fill.ForeColor.RGB = RGB(255, 255, 255)
                    End With


                End If
            Catch ex As Exception

            End Try



        End With


    End Sub

    'Public Sub defineLineAppearance(ByVal myproject As clsProjekt, ByVal myphase As clsPhase, ByVal lnumber As Integer, ByRef myShape As Excel.Shape, ByVal linienDicke As Double)
    '    'Dim pColor As Integer

    '    'With myphase

    '    '    pColor = CInt(.Farbe)

    '    'End With

    '    With myShape

    '        'With .Line
    '        '    .Visible = Microsoft.Office.Core.MsoTriState.msoTrue
    '        '    .ForeColor.RGB = pColor
    '        '    .Transparency = 0
    '        '    .Weight = CSng(linienDicke)
    '        'End With


    '        '.TextFrame2.TextRange.Text = ""
    '        'If lnumber > 0 And Not roentgenBlick.isOn Then

    '        '    With .TextFrame2
    '        '        .MarginLeft = 0
    '        '        .MarginRight = 0
    '        '        .MarginBottom = 0
    '        '        .MarginTop = 0
    '        '        .WordWrap = Microsoft.Office.Core.MsoTriState.msoFalse
    '        '        .VerticalAnchor = MsoVerticalAnchor.msoAnchorMiddle
    '        '        .HorizontalAnchor = MsoHorizontalAnchor.msoAnchorCenter
    '        '        .TextRange.Text = lnumber.ToString
    '        '        .TextRange.Font.Size = awinSettings.fontsizeLegend
    '        '        .TextRange.Font.Fill.ForeColor.RGB = RGB(255, 255, 255)
    '        '    End With


    '        'End If


    '    End With


    'End Sub


    Public Sub defineResultAppearance(ByVal myproject As clsProjekt, ByVal number As Integer, ByRef resultShape As Excel.Shape, ByVal bewertung As clsBewertung)
        'Dim pcolor As Object
        'Dim status As String

        'With myproject
        '    pcolor = .farbe
        '    status = .Status
        'End With




        With resultShape

            If awinSettings.mppShowAmpel = True Then
                .Glow.Radius = 5
                .Glow.Color.RGB = CInt(bewertung.color)
            End If

            'With .Line
            '    '.Visible = Microsoft.Office.Core.MsoTriState.msoTrue
            '    .Visible = MsoTriState.msoTrue
            '    .ForeColor.RGB = RGB(255, 255, 255)
            '    '.ForeColor.RGB = bewertung.color
            '    '.Transparency = 0
            'End With

            'With .Fill
            '    .ForeColor.RGB = CInt(bewertung.color)
            '    .ForeColor.TintAndShade = 0
            '    '.ForeColor.Brightness = 0.25
            '    .Transparency = 0.0

            '    'If roentgenBlick.isOn Then
            '    '    .Transparency = 0.8
            '    'Else
            '    '    If status = ProjektStatus(0) Then
            '    '        .Transparency = 0.35
            '    '    Else
            '    '        .Transparency = 0.0
            '    '    End If
            '    'End If

            '    .Solid()

            'End With


            Try
                If .TextFrame2.HasText <> MsoTriState.msoFalse Then
                    .TextFrame2.TextRange.Text = ""
                End If

                If number > 0 And Not roentgenBlick.isOn Then

                    With .TextFrame2
                        .MarginLeft = 0
                        .MarginRight = 0
                        .MarginBottom = 0
                        .MarginTop = 0
                        .WordWrap = Microsoft.Office.Core.MsoTriState.msoFalse
                        .VerticalAnchor = MsoVerticalAnchor.msoAnchorMiddle
                        .HorizontalAnchor = MsoHorizontalAnchor.msoAnchorCenter
                        .TextRange.Text = number.ToString
                        .TextRange.Font.Size = awinSettings.fontsizeLegend
                        .TextRange.Font.Fill.ForeColor.RGB = RGB(255, 255, 255)
                    End With


                End If

            Catch ex As Exception

            End Try


        End With

    End Sub

    Public Sub defineShapeAppearance(ByRef myproject As clsProjekt, ByRef projectShape As Excel.Shape)
        Dim pcolor As Object = XlRgbColor.rgbAqua
        Dim schriftFarbe As Long
        Dim schriftGroesse As Integer
        Dim status As String = ""
        Dim pMarge As Double
        Dim pname As String
        Dim diffToPrev As Boolean
        Dim ampel As Integer
        Dim showAmpel As Boolean = False
        Dim showResults As Boolean = True
        Dim myshape As Excel.Shape


        Try
            If projectShape.AlternativeText = CInt(PTshty.projektL).ToString Or _
                    projectShape.AutoShapeType = MsoAutoShapeType.msoShapeRoundedRectangle Then
                myshape = projectShape
            Else
                myshape = CType(projectShape.GroupItems.Item(1), Excel.Shape)
            End If
        Catch ex As Exception
            myshape = projectShape
        End Try


        Try
            With myproject
                pcolor = .farbe
                schriftFarbe = CLng(.Schriftfarbe)
                schriftGroesse = .Schrift
                status = .Status
                pMarge = .ProjectMarge
                pname = .name
                ampel = .ampelStatus
                diffToPrev = .diffToPrev
            End With
        Catch ex As Exception

        End Try


        With myshape

            Try
                If status = ProjektStatus(2) Or diffToPrev Then
                    ' beauftragt, aber noch nicht wieder freigegeben ... 

                    .Glow.Color.RGB = CInt(awinSettings.glowColor)
                    .Glow.Color.TintAndShade = 0
                    .Glow.Color.Brightness = 0
                    .Glow.Transparency = 0.4
                    .Glow.Radius = 10

                Else
                    .Glow.Color.RGB = RGB(255, 255, 255)
                    .Glow.Transparency = 1.0
                End If
            Catch ex As Exception

            End Try


            ' hier muss jetzt unterschieden werden, ob die Projektlinie gezeichnet wurde oder der Balken 

            If awinSettings.drawProjectLine Then

                Try
                    With .Line
                        .ForeColor.RGB = CInt(pcolor)
                        .Transparency = 0
                        .Weight = 4.0
                        .DashStyle = MsoLineDashStyle.msoLineDash
                    End With
                Catch ex As Exception

                End Try

                ' Darstellung, fixiert oder nicht fixiert 

                Try

                    With .Line
                        If status = ProjektStatus(0) Then
                            .BeginArrowheadStyle = MsoArrowheadStyle.msoArrowheadOval
                            .EndArrowheadStyle = MsoArrowheadStyle.msoArrowheadOval
                        Else
                            .BeginArrowheadStyle = MsoArrowheadStyle.msoArrowheadDiamond
                            .EndArrowheadStyle = MsoArrowheadStyle.msoArrowheadDiamond
                        End If

                    End With



                Catch ex As Exception

                End Try

            Else

                Try
                    With .Fill
                        '.Visible = msoTrue
                        .ForeColor.RGB = CInt(pcolor)
                        .ForeColor.TintAndShade = 0
                        .ForeColor.Brightness = -0.25

                        If roentgenBlick.isOn Then
                            .Transparency = 0.8
                        Else
                            .Transparency = 0.0
                        End If

                        .Solid()

                    End With
                Catch ex As Exception

                End Try


                Try
                    With .TextFrame2
                        .VerticalAnchor = MsoVerticalAnchor.msoAnchorMiddle
                        .HorizontalAnchor = MsoHorizontalAnchor.msoAnchorNone
                        .TextRange.Font.Size = schriftGroesse
                        .TextRange.Font.Fill.ForeColor.RGB = CInt(schriftFarbe)
                    End With

                    If roentgenBlick.isOn Then

                        .TextFrame2.TextRange.Text = ""


                    Else
                        ' Änderung 13.10.14 in den Namen soll jetzt der Varianten-Name aufgenommen werden, sofern es einen gibt 

                        .TextFrame2.TextRange.Text = myproject.getShapeText

                        ' Ende Änderung 13.10.14
                    End If
                Catch ex As Exception

                End Try

                ' nur verändern, wenn es auch veränderbar ist 
                Try

                    If .Adjustments.Count > 0 Then
                        If status = ProjektStatus(0) Then
                            .Adjustments.Item(1) = 0.5
                        Else
                            .Adjustments.Item(1) = 0.25
                        End If
                    End If

                Catch ex As Exception

                End Try






            End If

        End With


    End Sub

    ''' <summary>
    ''' definiert das Aussehen eines Shapes im Modus , wenn alles Shapes gezeichnet werden 
    ''' </summary>
    ''' <param name="myproject"></param>
    ''' <param name="projectShape"></param>
    ''' <param name="phasenIndex"></param>
    ''' <remarks></remarks>
    Public Sub defineShapeAppearance(ByVal myproject As clsProjekt, ByRef projectShape As Excel.Shape, ByVal phasenIndex As Integer)

        Dim projectColor As Object = Nothing, phaseColor As Object = RGB(255, 255, 255)
        Dim whiteColor As Object = RGB(255, 255, 255)
        Dim status As String = ""
        Dim pMarge As Double
        Dim pname As String
        Dim ampel As Integer
        Dim showAmpel As Boolean = False
        Dim showResults As Boolean = True
        Dim myshape As Excel.Shape
        Dim myphase As clsPhase


        Try
            myphase = myproject.getPhase(phasenIndex)


        Catch ex As Exception
            Throw New ArgumentException("Phase " & phasenIndex.ToString & _
                                        " existiert nicht ...")
        End Try

        Try

            myshape = CType(projectShape.GroupItems.Item(phasenIndex), Excel.Shape)

        Catch ex As Exception
            myshape = projectShape
        End Try

        Try
            With myproject
                projectColor = .farbe
                status = .Status
                pMarge = .ProjectMarge
                pname = .name
                ampel = .ampelStatus
            End With

        Catch ex As Exception

        End Try


        With myshape


            Try
                If status = ProjektStatus(2) Then

                    If phasenIndex = 1 Then
                        ' beauftragt, aber noch nicht wieder freigegeben ... 

                        .Glow.Color.RGB = CInt(awinSettings.glowColor)
                        .Glow.Color.TintAndShade = 0
                        .Glow.Color.Brightness = 0
                        .Glow.Transparency = 0.4
                        .Glow.Radius = 10
                    Else
                        .Glow.Color.RGB = RGB(255, 255, 255)
                        .Glow.Transparency = 1.0
                    End If
                Else
                    .Glow.Color.RGB = RGB(255, 255, 255)
                    .Glow.Transparency = 1.0
                End If
            Catch ex As Exception

            End Try


            Try
                With .Line
                    '.Visible = Microsoft.Office.Core.MsoTriState.msoTrue
                    '.Visible = Microsoft.Office.Core.MsoTriState.msoFalse
                    'If pMarge < 0 Then
                    '    .ForeColor.RGB = RGB(255, 0, 0)
                    '    .Weight = 2.0
                    'Else
                    '    .ForeColor.RGB = pcolor
                    'End If
                    If phasenIndex = 1 Then
                        .ForeColor.RGB = CInt(projectColor)
                        .Transparency = 0
                    Else
                        '.ForeColor.RGB = CInt(phaseColor)
                        .Transparency = 0
                    End If

                End With
            Catch ex As Exception

            End Try

            Try
                With .Fill

                    ' geändert wegen Änder
                    If phasenIndex = 1 Then
                        .ForeColor.RGB = CInt(projectColor)
                        'Else
                        '    .ForeColor.RGB = CInt(phaseColor)
                    End If

                    '.ForeColor.TintAndShade = 0
                    '.ForeColor.Brightness = -0.25

                    If roentgenBlick.isOn Then
                        .Transparency = 0.8
                    Else
                        .Transparency = 0.0
                    End If

                    If phasenIndex = 1 Then
                        .Solid()
                    End If


                End With

            Catch ex As Exception

            End Try


            Try

                If phasenIndex = 1 Then
                    If roentgenBlick.isOn Then

                        .TextFrame2.TextRange.Text = ""

                    Else

                        With .TextFrame2
                            .VerticalAnchor = MsoVerticalAnchor.msoAnchorMiddle
                            .HorizontalAnchor = MsoHorizontalAnchor.msoAnchorNone
                            .WordWrap = MsoTriState.msoFalse
                        End With

                        .TextFrame2.TextRange.Text = myproject.getShapeText


                    End If
                End If


            Catch ex As Exception

            End Try


            Try
                If .Adjustments.Count > 0 Then

                    If status = ProjektStatus(0) Then
                        .Adjustments.Item(1) = 0.5
                    Else
                        .Adjustments.Item(1) = 0.25
                    End If

                End If
            Catch ex As Exception

            End Try



        End With


    End Sub

    ' ''' <summary>
    ' ''' passt die Shape Darstellung dem veränderten Projekt pname an  
    ' ''' </summary>
    ' ''' <remarks></remarks>
    'Public Sub updateShapeinPlantafel(ByVal pname As String)
    '    Dim eeWasTrue As Boolean = False
    '    Dim suWasTrue As Boolean = False
    '    Dim zeile As Integer, spalte As Integer
    '    Dim laenge As Integer
    '    Dim status As String
    '    Dim top As Double, left As Double, width As Double, height As Double
    '    Dim magicBoardShapes As Excel.Shapes = appInstance.Worksheets(arrWsNames(3)).shapes
    '    Dim shpelement As Excel.Shape



    '    Dim hproj As New clsProjekt


    '    ' bestimmen der X- bzw. Y Position in der Plantafel 
    '    Try
    '        hproj = ShowProjekte.getProject(pname)
    '        With hproj
    '            zeile = .tfZeile
    '            spalte = .tfspalte
    '            laenge = .Dauer
    '            status = .Status
    '        End With
    '    Catch ex As Exception
    '        Call MsgBox("Fehler in clearProjektinPlantafel (Auslesen XPos, YPos, Dauer) von " & pname)
    '        Exit Sub
    '    End Try

    '    '
    '    ' hier wird in Plan Tafel das entsprechende Shape von der Erscheinung angepasst, ggf auch auf eine neue Zeile gesetzt ... 
    '    '
    '    Dim formerEE As Boolean = appInstance.EnableEvents
    '    Dim formerSU As Boolean = appInstance.ScreenUpdating
    '    appInstance.EnableEvents = False
    '    appInstance.ScreenUpdating = False

    '    With appInstance.Worksheets(arrWsNames(3))
    '        Try
    '            shpelement = magicBoardShapes.Item(pname)
    '            Dim myCollection As New Collection
    '            myCollection.Add(pname)
    '            zeile = findeMagicBoardPosition(myCollection, pname, zeile, spalte, laenge)

    '            ' jetzt ist eine passende Position gefunden ... die zugehörigen Shape Koordinaten werden berechnet 
    '            With hproj
    '                .tfZeile = zeile
    '                .CalculateShapeCoord(top, left, width, height)
    '            End With

    '            With shpelement
    '                .Top = top
    '                .Left = left
    '                .Width = width
    '                .Height = height
    '            End With

    '        Catch ex As Exception
    '            appInstance.EnableEvents = formerEE
    '            appInstance.ScreenUpdating = formerSU
    '            Throw New ArgumentException("updateProjektinPlantafel: kein Shape für Projekt " & pname & " gefunden")
    '        End Try


    '    End With

    '    appInstance.EnableEvents = formerEE
    '    appInstance.ScreenUpdating = formerSU


    'End Sub


    ''' <summary>
    ''' löscht die zeicherische Darstellung des Projektes auf der Plantafel 
    ''' </summary>
    ''' <remarks></remarks>
    Public Sub clearProjektinPlantafel(ByVal pname As String)
        Dim eeWasTrue As Boolean = False
        Dim suWasTrue As Boolean = False
        'Dim XPos As Integer, YPos As Integer
        'Dim laenge As Integer
        'Dim tmpshapes As Excel.Shapes = appInstance.ActiveSheet.shapes
        Dim tmpshapes As Excel.Shapes = CType(appInstance.Worksheets(arrWsNames(3)), Excel.Worksheet).Shapes
        Dim shpelement As Excel.Shape

        Dim formerEE As Boolean = appInstance.EnableEvents
        appInstance.EnableEvents = False


        'Dim formerSU As Boolean = appInstance.ScreenUpdating
        'appInstance.ScreenUpdating = False

        ' Lösche das Shape Element
        Try
            shpelement = tmpshapes.Item(pname)
            With shpelement
                projectboardShapes.remove(shpelement)
            End With
        Catch ex As Exception

        End Try

        Try
            Dim hproj As clsProjekt = ShowProjekte.getProject(pname)
            Dim shpuid As String = hproj.shpUID
            hproj.shpUID = ""
            ShowProjekte.shpListe.Remove(shpuid)
        Catch ex As Exception

        End Try

        ' Änderung 26.7.13 
        If roentgenBlick.isOn Then
            Call NoshowNeedsofProject(pname)
        End If



        appInstance.EnableEvents = formerEE
        'appInstance.ScreenUpdating = formerSU


    End Sub
    ''' <summary>
    ''' zeichnet den Pfeil, der anzeigt, um wieviel ein Projekt bei Optimierung verschoben werden würde
    ''' </summary>
    ''' <param name="pname"></param>
    ''' <remarks></remarks>
    Public Sub ZeichneMoveLineOfProjekt(ByRef pname As String)

        Dim start As Integer
        Dim laenge As Integer
        Dim pcolor As Integer, schriftfarbe As Object, fillColor As Integer, borderColor As Integer
        Dim schriftgroesse As Integer
        Dim zeilenOffset As Integer = 1
        Dim spaltenOffset As Integer = 0
        Dim hproj As clsProjekt
        Dim leftDrawn As Boolean
        Dim moveLength As Integer
        Dim tfz As Integer, tfs As Integer
        Dim top As Double, left As Double, width As Double, height As Double

        Dim straightLine As MsoConnectorType = MsoConnectorType.msoConnectorStraight


        hproj = ShowProjekte.getProject(pname)
        With hproj
            laenge = .anzahlRasterElemente
            start = .Start + .StartOffset
            moveLength = .StartOffset
            pcolor = CInt(.farbe)
            schriftfarbe = .Schriftfarbe
            schriftgroesse = .Schrift
            tfz = .tfZeile
            tfs = .tfspalte
        End With

        Dim formerEE As Boolean = appInstance.EnableEvents
        appInstance.EnableEvents = False


        If moveLength <> 0 Then
            height = 0.4 * boxHeight

            If moveLength < 0 Then
                leftDrawn = True
                left = (tfs - 1 + 0.5) * boxWidth + moveLength * boxWidth ' movelength ist negativ , deshalb "+"
                width = moveLength * boxWidth * (-1)
                top = topOfMagicBoard + (tfz - 1 + 0.75) * boxHeight
            Else
                leftDrawn = False
                left = (tfs + laenge - 1 - 0.5) * boxWidth
                top = topOfMagicBoard + (tfz - 1 + 0.25) * boxHeight
                width = moveLength * boxWidth
            End If

            Dim shp As Excel.Shape
            With appInstance.Worksheets(arrWsNames(3))


                fillColor = RGB(255, 255, 255)
                borderColor = pcolor

                If leftDrawn Then
                    shp = CType(.Shapes, Excel.Shapes).AddShape(Type:=Microsoft.Office.Core.MsoAutoShapeType.msoShapeLeftArrow, _
                                Left:=CSng(left), Top:=CSng(top), Width:=CSng(width), Height:=CSng(height))

                Else
                    shp = CType(.Shapes, Excel.Shapes).AddShape(Type:=Microsoft.Office.Core.MsoAutoShapeType.msoShapeRightArrow, _
                                Left:=CSng(left), Top:=CSng(top), Width:=CSng(width), Height:=CSng(height))
                End If

                ' jetzt wird der Pfeil gezeichnet




                With shp
                    With .Fill
                        .ForeColor.RGB = fillColor
                        .Transparency = 0.0
                    End With
                    With .Line
                        '.Visible = True
                        .Weight = 1.5
                        .ForeColor.RGB = borderColor
                        .Transparency = 0
                    End With

                End With



            End With



        End If

        appInstance.EnableEvents = formerEE

    End Sub

    Public Sub aktualisiereZeilenNrInProjekt(ByVal zeile As Integer, ByVal anzahl As Integer)

        For Each kvp As KeyValuePair(Of String, clsProjekt) In ShowProjekte.Liste
            With kvp.Value
                If .tfZeile >= zeile Then
                    .tfZeile = .tfZeile + anzahl
                End If
            End With
        Next kvp

    End Sub

    Public Sub visualisiereTeilErgebnis(ByVal pname As String)

        appInstance.ScreenUpdating = False
        Call diagramsVisible(False)
        'Call ZeichneProjektinPlanTafel(pname)
        Call ZeichneMoveLineOfProjekt(pname)
        Call awinNeuZeichnenDiagramme(1)
        Call diagramsVisible(True)
        appInstance.ScreenUpdating = True

    End Sub


    Public Sub visualisiereErgebnis()

        appInstance.ScreenUpdating = False
        Call diagramsVisible(False)
        'Call awinZeichnePlanTafel()
        Call awinNeuZeichnenDiagramme(1)
        Call diagramsVisible(True)
        appInstance.ScreenUpdating = True
    End Sub


    Public Sub insertZeile(ByVal zeilenNr As Integer)

        With appInstance.ActiveSheet
            .Rows(zeilenNr).Insert()
            .Rows(zeilenNr).Interior.ColorIndex = -4142
            .Rows(900).delete()
            Call aktualisiereZeilenNrInProjekt(zeilenNr, 1)
        End With
    End Sub

    ''' <summary>
    ''' selektiert in der Projekt Tafel das Projekt mit Namen pname
    ''' </summary>
    ''' <param name="pname"></param>
    ''' <remarks></remarks>
    Public Sub awinSelectProjectiST(ByRef pname As String, ByVal calledFromPf As Boolean)
        Dim hproj As clsProjekt
        Dim tfz As Integer, tfs As Integer
        Dim projektfarbe As Object
        Dim schriftfarbe As Object
        Dim realname As String = ""
        Dim tmparray() As String
        Dim allShapes As Excel.Shapes
        Dim selShape As Excel.Shape
        Dim shpUID As String


        Try
            allShapes = CType(appInstance.ActiveSheet, Excel.Worksheet).Shapes
        Catch ex As Exception
            allShapes = Nothing
        End Try

        If Not allShapes Is Nothing Then


            If calledFromPf Then
                ' der Name muss jetzt um das (xy.z%) bereinigt werden
                tmparray = pname.Split(New Char() {CChar("(")}, 10)
                Dim i As Integer
                For i = 0 To UBound(tmparray) - 1
                    realname = realname & tmparray(i)
                Next
                pname = realname.Trim
            End If

            If ShowProjekte.contains(pname) Then
                hproj = ShowProjekte.getProject(pname)
                With hproj
                    shpUID = .shpUID
                    'projektLaenge = .Dauer
                    projektfarbe = .farbe
                    schriftfarbe = .Schriftfarbe
                    tfz = .tfZeile
                    tfs = .tfspalte
                End With


                appInstance.EnableEvents = False

                Try
                    selShape = allShapes.Item(pname)
                    selShape.Select()
                Catch ex As Exception

                End Try



                appInstance.EnableEvents = True



            Else
                Call MsgBox("Projekt " & pname & " wurde nicht gefunden")
            End If


        End If
    End Sub

    Public Sub awinSelectProjectiPF(ByRef pname As String)
        Dim chtobj As ChartObject, rightObject As ChartObject = Nothing
        Dim found As Boolean = False
        Dim diagramTitle As String = "strategischer Fit, Risiko & Marge"
        Dim ptNr As Integer
        Dim chartPT As Point
        Dim anzPts As Integer

        ' das kann hier auskommentiert werden, da ab jetzt die Labels immer angezeigt werden 
        ' evtl wird es später wieder reingenommen ... 


        Try
            anzPts = UBound(PfChartBubbleNames)
        Catch ex As Exception
            Exit Sub
        End Try

        If anzPts >= 0 Then
            With CType(appInstance.Worksheets(arrWsNames(3)), Excel.Worksheet)
                For Each chtobj In CType(.ChartObjects, Excel.ChartObjects)
                    If chtobj.Chart.ChartTitle.Text = diagramTitle Then
                        found = True
                        rightObject = chtobj
                    End If
                Next
            End With

            ' jetzt muss bestimmt werden, die "wievielte Bubble " pname ist ...
            If found Then
                ptNr = 1
                While PfChartBubbleNames(ptNr - 1) <> pname And ptNr <= anzPts
                    ptNr = ptNr + 1
                End While

                If ptNr <= anzPts Then

                    Dim formerUpdate As Boolean = appInstance.ScreenUpdating
                    appInstance.ScreenUpdating = False

                    Try
                        With rightObject.Chart.SeriesCollection(1)
                            .ApplyDataLabels(Type:=Excel.XlDataLabelsType.xlDataLabelsShowNone)
                            chartPT = CType(.points(ptNr), Excel.Point)
                            chartPT.ApplyDataLabels(Type:=Excel.XlDataLabelsType.xlDataLabelsShowLabel)
                            chartPT.DataLabel.Text = pname
                        End With
                    Catch ex As Exception

                    End Try


                    appInstance.ScreenUpdating = formerUpdate
                End If

            End If
        End If
    End Sub



    ' ''' <summary>
    ' ''' speichert alle Projekte, die aktuell in Show- bzw NoShow sind
    ' ''' </summary>
    ' ''' <remarks></remarks>
    'Public Sub awinExportAllProjects()
    '    '
    '    Dim dateinameQ As String, dateinameZ As String


    '    appInstance.ScreenUpdating = False
    '    Try
    '        ' hier muss jetzt das File Projekt Detail aufgemacht werden ...
    '        appInstance.Workbooks.Open(awinPath & projektAustausch)


    '        For Each kvp As KeyValuePair(Of String, clsProjekt) In AlleProjekte.liste

    '            Try

    '                Call awinExportProject(kvp.Value)

    '            Catch ex As Exception

    '            End Try


    '        Next kvp

    '        For Each kvp As KeyValuePair(Of String, clsProjekt) In DeletedProjekte.Liste
    '            dateinameQ = awinPath & projektFilesOrdner & "\" & kvp.Key & ".xlsm"
    '            dateinameZ = awinPath & deletedFilesOrdner & "\" & kvp.Key & ".xlsm"
    '            Try
    '                My.Computer.FileSystem.MoveFile(dateinameQ, dateinameZ, True)
    '            Catch ex As Exception

    '            End Try


    '        Next kvp

    '    Catch ex As Exception
    '        Call MsgBox(ex.Message)
    '        Throw New ArgumentException("Abbruch - es konnten nicht ale Projekte gesichert werden ...")
    '        Exit Sub
    '    End Try

    '    appInstance.ActiveWorkbook.Close(SaveChanges:=False)  'UR:06.05.2014 PROJEKTSteckbrief darf nicht geändert werden
    '    appInstance.ScreenUpdating = True

    'End Sub


    ''' <summary>
    ''' Exportiert die Daten eines Projektes in einen Projekt-Steckbrief, ohne Berücksichtigung der Hierarchie
    ''' </summary>
    ''' <param name="hproj"></param>
    ''' <remarks></remarks>
    Public Sub awinExportProject(hproj As clsProjekt)

        Dim fileName As String
        Dim rng As Excel.Range, destinationRange As Excel.Range
        Dim zeile As Integer, spalte As Integer
        Dim rowOffset As Integer, columnOffset As Integer
        Dim delimiter As String = "."

        Dim formerEE As Boolean = appInstance.EnableEvents
        appInstance.EnableEvents = False


        zeile = 1
        spalte = 1

        ' Dateiname des Projectfiles '
        ' ur: 14.01.2015: Dateiname gleich dem Shape-Namen einschließlich VariantenNamen

        fileName = hproj.getShapeText & ".xlsx"

        'ur: 13.01.2015:  aus "fileName" werden die illegale Sonderzeichen eliminiert
        fileName = cleanFileName(fileName)

        ' fileName wird nun ergänzt mit dem passenden Pfad
        'fileName = awinPath & projektFilesOrdner & "\" & fileName
        fileName = exportOrdnerNames(PTImpExp.visbo) & "\" & fileName


        ' -------------------------------------------------
        ' hier werden die einzelnen Stamm-Daten in das entsprechende File geschrieben 
        ' -------------------------------------------------
        Try
            With appInstance.ActiveWorkbook.Worksheets("Stammdaten")

                .Unprotect(Password:="x")       ' Blattschutz aufheben

                ' Projekt-Name

                .range("Projekt_Name").value = hproj.name
                .range("Projekt_Name").interior.color = hproj.farbe
                .range("Projekt_Name").font.size = hproj.Schrift
                .range("Projekt_Name").font.color = hproj.Schriftfarbe


                ' Start

                .range("StartDatum").value = hproj.startDate

                ' Ende

                .range("EndeDatum").value = hproj.startDate.AddDays(hproj.dauerInDays - 1)


                'Projektleiter

                .range("Projektleiter").value = hproj.leadPerson

                ' Budget

                .range("Budget").value = hproj.Erloes.ToString("#####.#")

                'Kurzbeschreibung'

                .range("ProjektBeschreibung").value = hproj.description

                ' Ampel-Farbe

                .range("Bewertung").interior.color = awinSettings.AmpelNichtBewertet
                .range("Bewertung").value = hproj.ampelStatus


                ' Ampel-Bewertung 

                .range("BewertgErläuterung").value = hproj.ampelErlaeuterung


                '' Blattschutz setzen
                '.Protect(Password:="x", UserInterfaceOnly:=True, DrawingObjects:=True, Contents:=True, Scenarios:=True)

            End With
        Catch ex As Exception
            '' Blattschutz setzen
            'appInstance.Protect(Password:="x", UserInterfaceOnly:=True, DrawingObjects:=True, Contents:=True, Scenarios:=True)

            appInstance.EnableEvents = formerEE
            Throw New ArgumentException("Fehler in awinExportProject, Schreiben Stammdaten")
        End Try

        ' --------------------------------------------------
        ' jetzt werden die Ressourcen Bedarfe weggeschrieben 

        ' --------------------------------------------------

        Try
            With CType(appInstance.ActiveWorkbook.Worksheets("Ressourcen"), Excel.Worksheet)

                Dim tbl As Excel.Range

                .Unprotect(Password:="x")       ' Blattschutz aufheben

                zeile = 1
                spalte = 1

                ' die Monate eintragen 

                If awinSettings.zeitEinheit = "PM" Then

                    ' Dim htxt As String = hproj.startDate.ToShortDateString

                    ' Zeilen und Spalten-Offset für die Zeitleiste herausfinden
                    tbl = .Range("Zeitleiste")
                    rowOffset = tbl.Row         ' Reihen-Offset für die Zeitleiste
                    columnOffset = tbl.Column   ' Spalten-Offset für die Zeitleiste

                    ' Monat und Jahreszahl in die ersten beiden Felder der Zeitleiste eintragen'
                    .Range("Zeitleiste").Cells(columnOffset).value = "= StartDatum"

                    .Range("Zeitleiste").Cells(columnOffset + 1).value = "= EDATUM(D" & rowOffset & ",1"
                    .Range("Zeitleiste").Cells(columnOffset + 2).value = "= EDATUM(E" & rowOffset & ",1"

                    ' die ersten beiden Felder der Zeitleiste formatieren
                    rng = .Range(.Cells(rowOffset, columnOffset + 1), .Cells(rowOffset, columnOffset + 2))
                    rng.NumberFormat = "mmm-yy"
                    ' Die restliche Zeitleiste  formatieren
                    'rng = .range(.cells(startZeile, spalte), .cells(endZeile, spalte))
                    destinationRange = .Range(.Cells(rowOffset, columnOffset + 1), .Cells(rowOffset, columnOffset + 200))
                    With destinationRange
                        .HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter
                        .VerticalAlignment = Excel.XlVAlign.xlVAlignBottom
                        .NumberFormat = "mmm-yy"
                        .WrapText = False
                        .Orientation = 90
                        .AddIndent = False
                        .IndentLevel = 0
                        .ShrinkToFit = False
                        .ReadingOrder = Excel.Constants.xlContext
                        .MergeCells = False
                        .ColumnWidth = 4
                    End With

                    ' die Zeitleiste mit den Monatsangaben automatisch befüllen
                    rng.AutoFill(Destination:=destinationRange, Type:=XlAutoFillType.xlFillDefault)


                ElseIf awinSettings.zeitEinheit = "PW" Then
                ElseIf awinSettings.zeitEinheit = "PT" Then

                End If

                ' hier über alle Phasen ... 
                Dim cphase As clsPhase
                Dim p As Integer
                Dim phasenFarbe As Object
                Dim values() As Double
                Dim ErgebnisListe As New Collection
                Dim anzahlItems As Integer
                Dim r As Integer
                Dim d As Integer
                Dim itemNameID As String
                Dim dimension As Integer

                ' evtl hier vorher prüfen, ob es eine Phase mit Name hproj.name oder hproj.vorlagenName gibt; wenn nein , 
                ' muss hier der Projektname mit farbiger Gesamtdauer stehen 

                rowOffset = 1
                columnOffset = 1

                If hproj.CountPhases = 0 Then
                    ' Projekt-Name eintragen, Dauer einfärben, 28.2. genaues Start- und Endedatum in Kommentar eintragen

                    .Range("Phasen_des_Projekts").Cells(rowOffset, columnOffset).value = hproj.name
                    .Range("Phasen_des_Projekts").Cells(rowOffset, columnOffset).Interior.Color = hproj.farbe
                    rng = CType(.Range("Zeitmatrix")(.Cells(rowOffset, columnOffset), .Cells(rowOffset, columnOffset + hproj.anzahlRasterElemente - 1)), Excel.Range)
                    rng.Interior.Color = hproj.farbe
                    .Cells(rowOffset, columnOffset).AddComment()
                    With .Cells(rowOffset, columnOffset).Comment
                        .Visible = False
                        .Text(Text:="Start:" & Chr(10) & hproj.startDate)
                        .Shape.ScaleHeight(0.45, Microsoft.Office.Core.MsoTriState.msoFalse)
                    End With
                    .Cells(rowOffset, columnOffset + hproj.anzahlRasterElemente - 1).AddComment()
                    With .Cells(rowOffset, columnOffset + hproj.anzahlRasterElemente - 1).Comment
                        .Visible = False
                        .Text(Text:="Ende:" & Chr(10) & hproj.endeDate)
                        .Shape.ScaleHeight(0.45, Microsoft.Office.Core.MsoTriState.msoFalse)
                    End With
                    rowOffset = rowOffset + 1
                End If

                For p = 1 To hproj.CountPhases
                    cphase = hproj.getPhase(p)

                    ' Phasen-Name eintragen, Dauer einfärben
                    itemNameID = cphase.nameID

                    Try
                        phasenFarbe = cphase.Farbe
                    Catch ex As Exception
                        phasenFarbe = hproj.farbe
                    End Try

                    If itemNameID = rootPhaseName Then
                        ' Projekt-Name eintragen, Dauer einfärben
                        .Range("Phasen_des_Projekts").Cells(rowOffset, columnOffset).value = elemNameOfElemID(rootPhaseName)
                        .Range("Phasen_des_Projekts").Cells(rowOffset, columnOffset).Interior.Color = hproj.farbe
                        For d = 1 To hproj.anzahlRasterElemente
                            .Range("Zeitmatrix").Cells(rowOffset, columnOffset + d - 1).Interior.Color = hproj.farbe
                        Next d
                        ' Startdatum in Kommentar eintragen
                        .Range("Zeitmatrix").Cells(rowOffset, columnOffset).AddComment()
                        With .Range("Zeitmatrix").Cells(rowOffset, columnOffset).Comment
                            .Visible = False
                            .Text(Text:="Start:" & Chr(10) & hproj.startDate)
                            .Shape.ScaleHeight(0.45, Microsoft.Office.Core.MsoTriState.msoFalse)
                        End With
                        .Range("Zeitmatrix").Cells(rowOffset, columnOffset + hproj.anzahlRasterElemente - 1).AddComment()
                        With .Range("Zeitmatrix").Cells(rowOffset, columnOffset + hproj.anzahlRasterElemente - 1).Comment
                            .Visible = False
                            .Text(Text:="Ende:" & Chr(10) & hproj.endeDate)
                            .Shape.ScaleHeight(0.45, Microsoft.Office.Core.MsoTriState.msoFalse)
                        End With


                        d = CInt(appInstance.WorksheetFunction.CountA(.Range("Phasen_des_Projekts")))

                    Else
                        .Range("Phasen_des_Projekts").Cells(rowOffset, columnOffset).value = elemNameOfElemID(itemNameID)
                        For d = 1 To cphase.relEnde - cphase.relStart + 1
                            .Range("Zeitmatrix").Cells(rowOffset, cphase.relStart + d - 1).Interior.Color = phasenFarbe
                        Next d

                        ' Kommentar mit Start- und Endedatum eintragen
                        If cphase.relStart = cphase.relEnde Then
                            ' cphase ist nur ein Kästchen breit, d.h. Start-und EndeDatum müssen in einem Kommentar stehen
                            .Range("Zeitmatrix").Cells(rowOffset, cphase.relStart).AddComment()
                            With .Range("Zeitmatrix").Cells(rowOffset, cphase.relStart).Comment
                                .Visible = False
                                .Text(Text:="Start:" & Chr(10) & cphase.getStartDate & Chr(10) & "Ende:" & Chr(10) & cphase.getEndDate)
                            End With
                        Else
                            .Range("Zeitmatrix").Cells(rowOffset, cphase.relStart).AddComment()
                            With .Range("Zeitmatrix").Cells(rowOffset, cphase.relStart).Comment
                                .Visible = False
                                .Text(Text:="Start:" & Chr(10) & cphase.getStartDate)
                                .Shape.ScaleHeight(0.45, Microsoft.Office.Core.MsoTriState.msoFalse)
                            End With
                            .Range("Zeitmatrix").Cells(rowOffset, cphase.relEnde).AddComment()
                            With .Range("Zeitmatrix").Cells(rowOffset, cphase.relEnde).Comment
                                .Visible = False
                                .Text(Text:="Ende:" & Chr(10) & cphase.getEndDate)
                                .Shape.ScaleHeight(0.45, Microsoft.Office.Core.MsoTriState.msoFalse)
                            End With
                        End If
                        ' ende Kommentar eintragen in Ressourcen
                    End If

                    rowOffset = rowOffset + 1



                    anzahlItems = cphase.countRoles

                    Dim itemName As String
                    ' jetzt werden Rollen geschrieben 
                    For r = 1 To anzahlItems
                        itemName = cphase.getRole(r).name
                        dimension = cphase.getRole(r).getDimension
                        'ReDim values(cphase.relEnde - cphase.relStart)
                        ReDim values(dimension)
                        values = cphase.getRole(r).Xwerte
                        .Range("RollenKosten_des_Projekts").Cells(rowOffset, columnOffset).value = itemName

                        For d = 1 To dimension + 1
                            .Range("Zeitmatrix").Cells(rowOffset, cphase.relStart + d - 1).Interior.Color = phasenFarbe
                            .Range("Zeitmatrix").Cells(rowOffset, cphase.relStart + d - 1).Value = values(d - 1)
                        Next d
                        rowOffset = rowOffset + 1
                    Next r


                    ' jetzt werden Kosten geschrieben 

                    anzahlItems = cphase.countCosts

                    For k = 1 To anzahlItems
                        itemName = cphase.getCost(k).name
                        dimension = cphase.getCost(k).getDimension
                        ReDim values(dimension)
                        values = cphase.getCost(k).Xwerte
                        .Range("RollenKosten_des_Projekts").Cells(rowOffset, columnOffset).value = itemName
                        For d = 1 To dimension + 1
                            .Range("Zeitmatrix").Cells(rowOffset, cphase.relStart + d - 1).Interior.Color = phasenFarbe
                            .Range("Zeitmatrix").Cells(rowOffset, cphase.relStart + d - 1).Value = values(d - 1)
                        Next d
                        rowOffset = rowOffset + 1
                    Next
                    rowOffset = rowOffset + 1
                Next p

                '' Blattschutz setzen
                '.Protect(Password:="x", UserInterfaceOnly:=True, DrawingObjects:=True, Contents:=True, Scenarios:=True)

            End With
        Catch ex As Exception
            ' Blattschutz setzen
            appInstance.ActiveWorkbook.Worksheets("Ressourcen").Protect(Password:="x", UserInterfaceOnly:=True, DrawingObjects:=True, Contents:=True, Scenarios:=True)

            appInstance.EnableEvents = formerEE
            Throw New ArgumentException("Fehler in awinExportProject, Schreiben Ressourcen")
        End Try

        ' --------------------------------------------------
        ' jetzt werden die Settings in das unsichtbare Worksheet ("Settings") geschrieben 
        '
        Try

            With appInstance.ActiveWorkbook.Worksheets("Settings")

                .Unprotect(Password:="x")       ' Blattschutz aufheben

                zeile = 3
                Dim startZeile As Integer, endZeile As Integer
                Dim startRollen As Integer, startKosten As Integer
                spalte = 1
                Dim anzZeilen As Integer = 0
                rng = CType(.Range(.Cells(zeile, spalte), .Cells(zeile + 2000, spalte + 120)), Excel.Range)
                rng.Clear()

                ' ----------------------------------------- 
                ' Schreiben der Projektvorlagen
                '
                .cells(zeile, spalte).value = "Project-Vorlagen"
                .cells(zeile, spalte).interior.color = RGB(180, 180, 180)

                zeile = zeile + 1
                startZeile = zeile

                For Each kvp As KeyValuePair(Of String, clsProjektvorlage) In Projektvorlagen.Liste
                    .cells(zeile, spalte).value = kvp.Key
                    zeile = zeile + 1
                Next
                endZeile = zeile - 1

                If endZeile >= startZeile Then

                    If endZeile = startZeile Then
                        rng = CType(.cells(startZeile, spalte), Excel.Range)
                    Else
                        rng = CType(.range(.cells(startZeile, spalte), .cells(endZeile, spalte)), Excel.Range)
                    End If

                    appInstance.ActiveWorkbook.Names.Add(Name:="ProjektVorlagen", RefersTo:=rng)

                End If

                ' Schreiben des Delimiters
                .cells(zeile, spalte).value = delimiter
                zeile = zeile + 1

                ' ----------------------------------------- 
                ' Schreiben der Phasen
                '
                .cells(zeile, spalte).value = "Phasen"
                .cells(zeile, spalte).interior.color = RGB(180, 180, 180)

                zeile = zeile + 1
                startZeile = zeile

                For i = 1 To PhaseDefinitions.Count
                    .cells(zeile, spalte).value = PhaseDefinitions.getPhaseDef(i).name
                    .cells(zeile, spalte + 1).interior.color = PhaseDefinitions.getPhaseDef(i).farbe
                    zeile = zeile + 1
                Next

                endZeile = zeile - 1

                If endZeile >= startZeile Then

                    If endZeile = startZeile Then
                        rng = CType(.cells(startZeile, spalte), Excel.Range)
                    Else
                        rng = CType(.range(.cells(startZeile, spalte), .cells(endZeile, spalte)), Excel.Range)
                    End If
                    appInstance.ActiveWorkbook.Names.Add(Name:="Phasen", RefersTo:=rng)

                End If

                ' Schreiben des Delimiters
                .cells(zeile, spalte).value = delimiter
                zeile = zeile + 1


                ' ----------------------------------------- 
                ' Schreiben der Rollen und Kostenarten
                '

                .cells(zeile, spalte).value = "Rollen/Kostenarten"
                .cells(zeile, spalte).interior.color = RGB(180, 180, 180)

                zeile = zeile + 1
                startZeile = zeile
                startRollen = zeile

                For i = 1 To RoleDefinitions.Count
                    .cells(zeile, spalte).value = RoleDefinitions.getRoledef(i).name
                    zeile = zeile + 1
                Next

                endZeile = zeile - 1

                If endZeile >= startZeile Then

                    If endZeile = startZeile Then
                        rng = CType(.cells(startZeile, spalte), Excel.Range)
                    Else
                        rng = CType(.range(.cells(startZeile, spalte), .cells(endZeile, spalte)), Excel.Range)
                    End If
                    appInstance.ActiveWorkbook.Names.Add(Name:="Rollen", RefersTo:=rng)

                End If


                startKosten = zeile

                For i = 1 To CostDefinitions.Count - 1
                    ' die Personalkosten sind die letzte Kostenart, wird nicht mit aufgenommen, da sie 
                    ' automatisch berücksichtigt wird 
                    .cells(zeile, spalte).value = CostDefinitions.getCostdef(i).name
                    zeile = zeile + 1
                Next

                endZeile = zeile - 1
                If endZeile >= startKosten Then
                    rng = CType(.range(.cells(startKosten, spalte), .cells(endZeile, spalte)), Excel.Range)
                    appInstance.ActiveWorkbook.Names.Add(Name:="Kosten", RefersTo:=rng)

                End If

                If endZeile >= startZeile Then
                    rng = CType(.range(.cells(startZeile, spalte), .cells(endZeile, spalte)), Excel.Range)
                    appInstance.ActiveWorkbook.Names.Add(Name:="Rollen_Kostenarten", RefersTo:=rng)

                End If

                ' Schreiben des Delimiters
                .cells(zeile, spalte).value = delimiter
                zeile = zeile + 1


                ' -------------------------------------------
                ' Schreiben der Ampel-Farben 
                '


                .cells(zeile, spalte).value = "Ampel-Farben"
                .cells(zeile, spalte).interior.color = RGB(180, 180, 180)

                zeile = zeile + 1
                startZeile = zeile
                For i = 0 To 3
                    Select Case i
                        Case 0
                            .cells(zeile, spalte).value = "Ampel nicht bewertet"
                            .cells(zeile, spalte).interior.color = awinSettings.AmpelNichtBewertet
                        Case 1
                            .cells(zeile, spalte).value = "Ampel Grün"
                            .cells(zeile, spalte).interior.color = awinSettings.AmpelGruen
                        Case 2
                            .cells(zeile, spalte).value = "Ampel Gelb"
                            .cells(zeile, spalte).interior.color = awinSettings.AmpelGelb
                        Case 3
                            .cells(zeile, spalte).value = "Ampel Rot"
                            .cells(zeile, spalte).interior.color = awinSettings.AmpelRot
                    End Select
                    zeile = zeile + 1
                Next

                endZeile = zeile - 1

                If endZeile >= startZeile Then
                    rng = CType(.range(.cells(startZeile, spalte), .cells(endZeile, spalte)), Excel.Range)
                    appInstance.ActiveWorkbook.Names.Add(Name:="AmpelFarben", RefersTo:=rng)

                End If

                ' Schreiben des Delimiters
                .cells(zeile, spalte).value = delimiter
                zeile = zeile + 1


                ' ----------------------------------------- 
                ' Unsichtbarmachen des Tabellenblattes
                '

                .Visible = Excel.XlSheetVisibility.xlSheetHidden

            End With

        Catch ex As Exception
            appInstance.EnableEvents = formerEE
            Throw New ArgumentException("Fehler in awinExportProject, Schreiben Settings")
        End Try

        ' ----------------------------------------------
        ' jetzt werden die Termine weggeschrieben ....

        Try
            With appInstance.ActiveWorkbook.Worksheets("Termine")


                .Unprotect(Password:="x")       ' Blattschutz aufheben

                zeile = 1
                spalte = 1

            End With
        Catch ex As Exception
            With appInstance.ActiveWorkbook.Worksheets.Add
                .name = "Termine"
                ' Tabelle ErgebnTabelle muss hier eigentlich erzeugt werden
                appInstance.EnableEvents = formerEE
                Throw New ArgumentException("Fehler in awinExportProject, Schreiben Termine, Worksheet Termine existiert nicht")
            End With
        End Try

        ' --------------------------------------
        ' Worksheet Termine existriert jetzt ...

        With CType(appInstance.ActiveWorkbook.Worksheets("Termine"), Excel.Worksheet)

            .Unprotect(Password:="x")       ' Blattschutz aufheben

            Dim cphase As New clsPhase(hproj)
            Dim phaseName As String
            Dim r As Integer
            Dim cResult As New clsMeilenstein(parent:=cphase)
            Dim cBewertung As clsBewertung
            Dim phaseStart As Date
            Dim phaseEnde As Date
            Dim tbl As Excel.Range
            Dim tablename As String

            tablename = .ListObjects("ErgebnTabelle").Name
            tbl = .ListObjects("ErgebnTabelle").Range
            rowOffset = tbl.Row
            columnOffset = tbl.Column


            For p = 1 To hproj.CountPhases
                cphase = hproj.getPhase(p)

                If awinSettings.zeitEinheit = "PM" Then
                    phaseStart = hproj.startDate.AddDays(cphase.startOffsetinDays)
                    phaseEnde = hproj.startDate.AddDays(cphase.startOffsetinDays + cphase.dauerInDays - 1)
                ElseIf awinSettings.zeitEinheit = "PW" Then
                    phaseStart = hproj.startDate.AddDays((cphase.relStart - 1) * 7)
                    phaseEnde = hproj.startDate.AddDays((cphase.relEnde - 1) * 7)
                ElseIf awinSettings.zeitEinheit = "PT" Then
                    phaseStart = hproj.startDate.AddDays(cphase.relStart - 1)
                    phaseEnde = hproj.startDate.AddDays(cphase.relEnde - 1)
                End If


                phaseName = cphase.name

                ' hier muss die Phase geschrieben werden
                .Cells(rowOffset + zeile, columnOffset).value = zeile
                .Cells(rowOffset + zeile, columnOffset + 1).value = phaseName
                .Cells(rowOffset + zeile, columnOffset + 2).value = ""
                .Cells(rowOffset + zeile, columnOffset + 3).value = phaseStart
                .Cells(rowOffset + zeile, columnOffset + 4).value = phaseEnde
                '.Cells(rowOffset + zeile, columnOffset + 5).value = " "
                .Cells(rowOffset + zeile, columnOffset + 5).value = "0"
                .Cells(rowOffset + zeile, columnOffset + 5).interior.color = awinSettings.AmpelNichtBewertet
                .Cells(rowOffset + zeile, columnOffset + 6).value = " "

                zeile = zeile + 1

                For r = 1 To cphase.countMilestones
                    cResult = cphase.getMilestone(r)

                    cBewertung = cResult.getBewertung(1)
                    'Try
                    '    cBewertung = cResult.getBewertung(1)
                    'Catch ex As Exception
                    '    cBewertung = New clsBewertung
                    'End Try
                    ' --------------------------------------------------------------------------------
                    ' Termine müssen in Tabelle eingetragen werden
                    '----------------------------------------------------------------------------------

                    .Cells(rowOffset + zeile, columnOffset).value = zeile
                    .Cells(rowOffset + zeile, columnOffset + 1).value = cResult.name
                    With CType(.Cells(rowOffset + zeile, columnOffset + 1), Excel.Range)
                        .IndentLevel = 2
                        .HorizontalAlignment = Excel.XlHAlign.xlHAlignLeft
                    End With
                    .Cells(rowOffset + zeile, columnOffset + 2).value = phaseName
                    .Cells(rowOffset + zeile, columnOffset + 3).value = ""
                    .Cells(rowOffset + zeile, columnOffset + 4).value = cResult.getDate
                    ' .Cells(rowOffset + zeile, columnOffset + 4).value = cResult.verantwortlich
                    .Cells(rowOffset + zeile, columnOffset + 5).value = cBewertung.colorIndex
                    .Cells(rowOffset + zeile, columnOffset + 5).interior.color = cBewertung.color
                    ' Zelle für Beschreibung in der Höhe anpassen, autom. Zeilenumbruch
                    .Cells(rowOffset + zeile, columnOffset + 6).value = cBewertung.description
                    .Cells(rowOffset + zeile, columnOffset + 6).Rows.Autofit()
                    .Cells(rowOffset + zeile, columnOffset + 6).WrapText = True

                    zeile = zeile + 1
                Next

            Next

            '' Blattschutz setzen
            '.Protect(Password:="x", UserInterfaceOnly:=True, DrawingObjects:=True, Contents:=True, Scenarios:=True)

        End With


        ' ----------------------------------------------
        ' jetzt werden die Attribute weggeschrieben ....

        Try
            With CType(appInstance.ActiveWorkbook.Worksheets("Attribute"), Excel.Worksheet)

                .Unprotect(Password:="x")       ' Blattschutz aufheben


                ' Projekt-Typ

                .Range("Projekt_Typ").Value = hproj.VorlagenName
                rng = .Range("Projekt_Typ")
                With rng
                    .HorizontalAlignment = Excel.XlHAlign.xlHAlignLeft
                    .IndentLevel = 1
                    .WrapText = False
                End With

                ' Status

                .Range("Status").Value = hproj.Status
                rng = .Range("Status")
                With rng
                    .HorizontalAlignment = Excel.XlHAlign.xlHAlignRight
                    .IndentLevel = 1
                    .WrapText = False
                End With

                ' Business_Unit

                .Range("Business_Unit").Value = hproj.businessUnit
                rng = .Range("Business_Unit")
                With rng
                    .HorizontalAlignment = Excel.XlHAlign.xlHAlignLeft
                    .IndentLevel = 1
                    .WrapText = False
                End With

                ' Strategischer Fit

                .Range("Strategischer_Fit").Value = hproj.StrategicFit
                rng = .Range("Strategischer_Fit")
                With rng
                    .HorizontalAlignment = Excel.XlHAlign.xlHAlignRight
                    .IndentLevel = 1
                    .WrapText = False
                End With

                ' Risiko

                .Range("Risiko").Value = hproj.Risiko
                rng = .Range("Risiko")
                With rng
                    .HorizontalAlignment = Excel.XlHAlign.xlHAlignRight
                    .IndentLevel = 1
                    .WrapText = False
                End With

                ' ur: 13.01.2015: Varianten_Name wird hier in das Tabellenblatt Attribute des Projekt-Steckbriefes eingetragen

                If Not IsNothing(hproj.variantName) And hproj.variantName <> "" Then

                    .Range("Variant_Name").Value = hproj.variantName
                    rng = .Range("Variant_Name")
                    With rng
                        .HorizontalAlignment = Excel.XlHAlign.xlHAlignLeft
                        .IndentLevel = 1
                        .WrapText = False
                    End With

                End If

                '' Blattschutz setzen
                '.Protect(Password:="x", UserInterfaceOnly:=True, DrawingObjects:=True, Contents:=True, Scenarios:=True)

            End With
        Catch ex As Exception
            '' Blattschutz setzen
            'appInstance.ActiveWorkbook.Worksheets("Attribute").Protect(Password:="x", UserInterfaceOnly:=True, DrawingObjects:=True, Contents:=True, Scenarios:=True)
            appInstance.EnableEvents = formerEE
            Throw New ArgumentException("Fehler in awinExportProject, Schreiben Attribute")
        End Try


        Try

            My.Computer.FileSystem.DeleteFile(fileName)
        Catch ex As Exception

        End Try

        Try

            appInstance.ActiveWorkbook.SaveAs(fileName, _
                                          ConflictResolution:=XlSaveConflictResolution.xlLocalSessionChanges
                                          )

        Catch ex As Exception
            appInstance.EnableEvents = formerEE
            Throw New ArgumentException("Fehler beim Datei-Schreiben")
        End Try


        appInstance.EnableEvents = formerEE


    End Sub

    ''' <summary>
    ''' Exportiert das Projekt hproj in einen Projektsteckbrief mit verwendeter Hierarchischem Aufbau der Phasen und Meilensteine
    ''' </summary>
    ''' <param name="hproj"></param>
    ''' <remarks></remarks>
    Public Sub awinExportProjectmitHrchy(hproj As clsProjekt)

        Dim fileName As String
        Dim rng As Excel.Range, destinationRange As Excel.Range
        Dim zeile As Integer, spalte As Integer
        Dim rowOffset As Integer, columnOffset As Integer
        Dim delimiter As String = "."
        Dim einrückJeStufe As String = "  "


        Dim formerEE As Boolean = appInstance.EnableEvents
        appInstance.EnableEvents = False


        zeile = 1
        spalte = 1

        ' Dateiname des Projectfiles '
        ' ur: 14.01.2015: Dateiname gleich dem Shape-Namen einschließlich VariantenNamen

        fileName = hproj.getShapeText & ".xlsx"

        'ur: 13.01.2015:  aus "fileName" werden die illegale Sonderzeichen eliminiert
        fileName = cleanFileName(fileName)

        ' fileName wird nun ergänzt mit dem passenden Pfad
        'fileName = awinPath & projektFilesOrdner & "\" & fileName
        fileName = exportOrdnerNames(PTImpExp.visbo) & "\" & fileName

        ' -------------------------------------------------
        ' hier werden die einzelnen Stamm-Daten in das entsprechende File geschrieben 
        ' -------------------------------------------------
        Try
            With appInstance.ActiveWorkbook.Worksheets("Stammdaten")

                .Unprotect(Password:="x")       ' Blattschutz aufheben

                ' Projekt-Name

                .range("Projekt_Name").value = hproj.name
                .range("Projekt_Name").interior.color = hproj.farbe
                .range("Projekt_Name").font.size = hproj.Schrift
                .range("Projekt_Name").font.color = hproj.Schriftfarbe


                ' Start

                .range("StartDatum").value = hproj.startDate

                ' Ende

                .range("EndeDatum").value = hproj.startDate.AddDays(hproj.dauerInDays - 1)


                'Projektleiter

                .range("Projektleiter").value = hproj.leadPerson

                ' Budget

                .range("Budget").value = hproj.Erloes.ToString("#####.#")

                'Kurzbeschreibung'

                .range("ProjektBeschreibung").value = hproj.description

                ' Ampel-Farbe

                .range("Bewertung").interior.color = awinSettings.AmpelNichtBewertet
                .range("Bewertung").value = hproj.ampelStatus


                ' Ampel-Bewertung 

                .range("BewertgErläuterung").value = hproj.ampelErlaeuterung


                '' Blattschutz setzen
                '.Protect(Password:="x", UserInterfaceOnly:=True, DrawingObjects:=True, Contents:=True, Scenarios:=True)

            End With
        Catch ex As Exception
            '' Blattschutz setzen
            'appInstance.Protect(Password:="x", UserInterfaceOnly:=True, DrawingObjects:=True, Contents:=True, Scenarios:=True)

            appInstance.EnableEvents = formerEE
            Throw New ArgumentException("Fehler in awinExportProject, Schreiben Stammdaten")
        End Try

        ' --------------------------------------------------
        ' jetzt werden die Ressourcen Bedarfe weggeschrieben 

        ' --------------------------------------------------

        Try
            With CType(appInstance.ActiveWorkbook.Worksheets("Ressourcen"), Excel.Worksheet)

                Dim tbl As Excel.Range

                .Unprotect(Password:="x")       ' Blattschutz aufheben

                zeile = 1
                spalte = 1

                ' die Monate eintragen 

                If awinSettings.zeitEinheit = "PM" Then

                    ' Dim htxt As String = hproj.startDate.ToShortDateString

                    ' Zeilen und Spalten-Offset für die Zeitleiste herausfinden
                    tbl = .Range("Zeitleiste")
                    rowOffset = tbl.Row         ' Reihen-Offset für die Zeitleiste
                    columnOffset = tbl.Column   ' Spalten-Offset für die Zeitleiste

                    ' Monat und Jahreszahl in die ersten beiden Felder der Zeitleiste eintragen'
                    .Range("Zeitleiste").Cells(columnOffset).value = "= StartDatum"

                    .Range("Zeitleiste").Cells(columnOffset + 1).value = "= EDATUM(D" & rowOffset & ",1"
                    .Range("Zeitleiste").Cells(columnOffset + 2).value = "= EDATUM(E" & rowOffset & ",1"

                    ' die ersten beiden Felder der Zeitleiste formatieren
                    rng = .Range(.Cells(rowOffset, columnOffset + 1), .Cells(rowOffset, columnOffset + 2))
                    rng.NumberFormat = "mmm-yy"
                    ' Die restliche Zeitleiste  formatieren
                    'rng = .range(.cells(startZeile, spalte), .cells(endZeile, spalte))
                    destinationRange = .Range(.Cells(rowOffset, columnOffset + 1), .Cells(rowOffset, columnOffset + 200))
                    With destinationRange
                        .HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter
                        .VerticalAlignment = Excel.XlVAlign.xlVAlignBottom
                        .NumberFormat = "mmm-yy"
                        .WrapText = False
                        .Orientation = 90
                        .AddIndent = False
                        .IndentLevel = 0
                        .ShrinkToFit = False
                        .ReadingOrder = Excel.Constants.xlContext
                        .MergeCells = False
                        .ColumnWidth = 4
                    End With

                    ' die Zeitleiste mit den Monatsangaben automatisch befüllen
                    rng.AutoFill(Destination:=destinationRange, Type:=XlAutoFillType.xlFillDefault)


                ElseIf awinSettings.zeitEinheit = "PW" Then
                ElseIf awinSettings.zeitEinheit = "PT" Then

                End If

                ' hier über alle Phasen ... 
                Dim cphase As clsPhase
                Dim p As Integer
                Dim phasenFarbe As Object
                Dim values() As Double
                Dim ErgebnisListe As New Collection
                Dim anzahlItems As Integer
                Dim r As Integer
                Dim d As Integer
                Dim itemNameID As String
                Dim dimension As Integer

                ' evtl hier vorher prüfen, ob es eine Phase mit Name hproj.name oder hproj.vorlagenName gibt; wenn nein , 
                ' muss hier der Projektname mit farbiger Gesamtdauer stehen 

                rowOffset = 1
                columnOffset = 1

                If hproj.CountPhases = 0 Then
                    ' Projekt-Name eintragen, Dauer einfärben, 28.2. genaues Start- und Endedatum in Kommentar eintragen

                    .Range("Phasen_des_Projekts").Cells(rowOffset, columnOffset).value = hproj.name
                    .Range("Phasen_des_Projekts").Cells(rowOffset, columnOffset).Interior.Color = hproj.farbe
                    rng = CType(.Range("Zeitmatrix")(.Cells(rowOffset, columnOffset), .Cells(rowOffset, columnOffset + hproj.anzahlRasterElemente - 1)), Excel.Range)
                    rng.Interior.Color = hproj.farbe
                    .Cells(rowOffset, columnOffset).AddComment()
                    With .Cells(rowOffset, columnOffset).Comment
                        .Visible = False
                        .Text(Text:="Start:" & Chr(10) & hproj.startDate)
                        .Shape.ScaleHeight(0.45, Microsoft.Office.Core.MsoTriState.msoFalse)
                    End With
                    .Cells(rowOffset, columnOffset + hproj.anzahlRasterElemente - 1).AddComment()
                    With .Cells(rowOffset, columnOffset + hproj.anzahlRasterElemente - 1).Comment
                        .Visible = False
                        .Text(Text:="Ende:" & Chr(10) & hproj.endeDate)
                        .Shape.ScaleHeight(0.45, Microsoft.Office.Core.MsoTriState.msoFalse)
                    End With
                    rowOffset = rowOffset + 1
                End If

                For p = 1 To hproj.CountPhases
                    cphase = hproj.getPhase(p)

                    ' Phasen-Name eintragen, Dauer einfärben
                    itemNameID = cphase.nameID

                    Try
                        phasenFarbe = cphase.Farbe
                    Catch ex As Exception
                        phasenFarbe = hproj.farbe
                    End Try

                    If itemNameID = rootPhaseName Then  ' rootPhaseName = "0§.§" als Konstante definiert

                        ' Projekt-Name eintragen, Dauer einfärben

                        .Range("Phasen_des_Projekts").Cells(rowOffset, columnOffset).value = elemNameOfElemID(rootPhaseName)
                        .Range("Phasen_des_Projekts").Cells(rowOffset, columnOffset).Interior.Color = hproj.farbe
                        For d = 1 To hproj.anzahlRasterElemente
                            .Range("Zeitmatrix").Cells(rowOffset, columnOffset + d - 1).Interior.Color = hproj.farbe
                        Next d
                        ' Startdatum in Kommentar eintragen
                        .Range("Zeitmatrix").Cells(rowOffset, columnOffset).AddComment()
                        With .Range("Zeitmatrix").Cells(rowOffset, columnOffset).Comment
                            .Visible = False
                            .Text(Text:="Start:" & Chr(10) & hproj.startDate)
                            .Shape.ScaleHeight(0.45, Microsoft.Office.Core.MsoTriState.msoFalse)
                        End With
                        .Range("Zeitmatrix").Cells(rowOffset, columnOffset + hproj.anzahlRasterElemente - 1).AddComment()
                        With .Range("Zeitmatrix").Cells(rowOffset, columnOffset + hproj.anzahlRasterElemente - 1).Comment
                            .Visible = False
                            .Text(Text:="Ende:" & Chr(10) & hproj.endeDate)
                            .Shape.ScaleHeight(0.45, Microsoft.Office.Core.MsoTriState.msoFalse)
                        End With


                        d = CInt(appInstance.WorksheetFunction.CountA(.Range("Phasen_des_Projekts")))

                    Else
                        ' ur:06.05.2015: hier müssen die Einrückungen erfolgen

                        Dim indlevel As Integer = hproj.hierarchy.getIndentLevel(itemNameID)
                        Dim phstr As String = ""

                        .Range("Phasen_des_Projekts").Cells(rowOffset, columnOffset).value = elemNameOfElemID(itemNameID)
                        With CType(.Range("Phasen_des_Projekts").Cells(rowOffset, columnOffset), Excel.Range)
                            .IndentLevel = indlevel * einrückTiefe
                            .HorizontalAlignment = Excel.XlHAlign.xlHAlignLeft
                        End With

                        For d = 1 To cphase.relEnde - cphase.relStart + 1
                            .Range("Zeitmatrix").Cells(rowOffset, cphase.relStart + d - 1).Interior.Color = phasenFarbe
                        Next d

                        ' Kommentar mit Start- und Endedatum eintragen
                        If cphase.relStart = cphase.relEnde Then
                            ' cphase ist nur ein Kästchen breit, d.h. Start-und EndeDatum müssen in einem Kommentar stehen
                            .Range("Zeitmatrix").Cells(rowOffset, cphase.relStart).AddComment()
                            With .Range("Zeitmatrix").Cells(rowOffset, cphase.relStart).Comment
                                .Visible = False
                                .Text(Text:="Start:" & Chr(10) & cphase.getStartDate & Chr(10) & "Ende:" & Chr(10) & cphase.getEndDate)
                            End With
                        Else
                            .Range("Zeitmatrix").Cells(rowOffset, cphase.relStart).AddComment()
                            With .Range("Zeitmatrix").Cells(rowOffset, cphase.relStart).Comment
                                .Visible = False
                                .Text(Text:="Start:" & Chr(10) & cphase.getStartDate)
                                .Shape.ScaleHeight(0.45, Microsoft.Office.Core.MsoTriState.msoFalse)
                            End With
                            .Range("Zeitmatrix").Cells(rowOffset, cphase.relEnde).AddComment()
                            With .Range("Zeitmatrix").Cells(rowOffset, cphase.relEnde).Comment
                                .Visible = False
                                .Text(Text:="Ende:" & Chr(10) & cphase.getEndDate)
                                .Shape.ScaleHeight(0.45, Microsoft.Office.Core.MsoTriState.msoFalse)
                            End With
                        End If
                        ' ende Kommentar eintragen in Ressourcen
                    End If

                    rowOffset = rowOffset + 1



                    anzahlItems = cphase.countRoles

                    Dim itemName As String
                    ' jetzt werden Rollen geschrieben 
                    For r = 1 To anzahlItems
                        itemName = cphase.getRole(r).name
                        dimension = cphase.getRole(r).getDimension
                        'ReDim values(cphase.relEnde - cphase.relStart)
                        ReDim values(dimension)
                        values = cphase.getRole(r).Xwerte
                        .Range("RollenKosten_des_Projekts").Cells(rowOffset, columnOffset).value = itemName

                        For d = 1 To dimension + 1
                            .Range("Zeitmatrix").Cells(rowOffset, cphase.relStart + d - 1).Interior.Color = phasenFarbe
                            .Range("Zeitmatrix").Cells(rowOffset, cphase.relStart + d - 1).Value = values(d - 1)
                        Next d
                        rowOffset = rowOffset + 1
                    Next r


                    ' jetzt werden Kosten geschrieben 

                    anzahlItems = cphase.countCosts

                    For k = 1 To anzahlItems
                        itemName = cphase.getCost(k).name
                        dimension = cphase.getCost(k).getDimension
                        ReDim values(dimension)
                        values = cphase.getCost(k).Xwerte
                        .Range("RollenKosten_des_Projekts").Cells(rowOffset, columnOffset).value = itemName
                        For d = 1 To dimension + 1
                            .Range("Zeitmatrix").Cells(rowOffset, cphase.relStart + d - 1).Interior.Color = phasenFarbe
                            .Range("Zeitmatrix").Cells(rowOffset, cphase.relStart + d - 1).Value = values(d - 1)
                        Next d
                        rowOffset = rowOffset + 1
                    Next
                    rowOffset = rowOffset + 1
                Next p

                '' Blattschutz setzen
                '.Protect(Password:="x", UserInterfaceOnly:=True, DrawingObjects:=True, Contents:=True, Scenarios:=True)

            End With
        Catch ex As Exception
            ' Blattschutz setzen
            appInstance.ActiveWorkbook.Worksheets("Ressourcen").Protect(Password:="x", UserInterfaceOnly:=True, DrawingObjects:=True, Contents:=True, Scenarios:=True)

            appInstance.EnableEvents = formerEE
            Throw New ArgumentException("Fehler in awinExportProject, Schreiben Ressourcen")
        End Try

        ' --------------------------------------------------
        ' jetzt werden die Settings in das unsichtbare Worksheet ("Settings") geschrieben 
        '
        Try

            With appInstance.ActiveWorkbook.Worksheets("Settings")

                .Unprotect(Password:="x")       ' Blattschutz aufheben

                zeile = 3
                Dim startZeile As Integer, endZeile As Integer
                Dim startMeilensteine As Integer
                Dim startRollen As Integer, startKosten As Integer
                spalte = 1
                Dim anzZeilen As Integer = 0
                rng = CType(.Range(.Cells(zeile, spalte), .Cells(zeile + 2000, spalte + 120)), Excel.Range)
                rng.Clear()

                ' ----------------------------------------- 
                ' Schreiben der Projektvorlagen
                '
                .cells(zeile, spalte).value = "Project-Vorlagen"
                .cells(zeile, spalte).interior.color = RGB(180, 180, 180)

                zeile = zeile + 1
                startZeile = zeile

                For Each kvp As KeyValuePair(Of String, clsProjektvorlage) In Projektvorlagen.Liste
                    .cells(zeile, spalte).value = kvp.Key
                    zeile = zeile + 1
                Next
                endZeile = zeile - 1

                If endZeile >= startZeile Then

                    If endZeile = startZeile Then
                        rng = CType(.cells(startZeile, spalte), Excel.Range)
                    Else
                        rng = CType(.range(.cells(startZeile, spalte), .cells(endZeile, spalte)), Excel.Range)
                    End If

                    appInstance.ActiveWorkbook.Names.Add(Name:="ProjektVorlagen", RefersTo:=rng)

                End If

                ' Schreiben des Delimiters
                .cells(zeile, spalte).value = delimiter
                zeile = zeile + 1

                ' ----------------------------------------- 
                ' Schreiben der Phasen
                '
                .cells(zeile, spalte).value = "Phasen"
                .cells(zeile, spalte).interior.color = RGB(180, 180, 180)

                zeile = zeile + 1
                startZeile = zeile

                For i = 1 To PhaseDefinitions.Count

                    Dim cphaseDef As clsPhasenDefinition = PhaseDefinitions.getPhaseDef(i)
                    .cells(zeile, spalte).value = cphaseDef.name
                    .cells(zeile, spalte + 1).interior.color = cphaseDef.farbe
                    zeile = zeile + 1
                Next

                endZeile = zeile - 1

                If endZeile >= startZeile Then

                    If endZeile = startZeile Then
                        rng = CType(.cells(startZeile, spalte), Excel.Range)
                    Else
                        rng = CType(.range(.cells(startZeile, spalte), .cells(endZeile, spalte)), Excel.Range)
                    End If
                    appInstance.ActiveWorkbook.Names.Add(Name:="Phasen", RefersTo:=rng)

                End If

                ' ----------------------------------------- 
                ' Schreiben der Meilensteine
                '
                .cells(zeile, spalte + 2).value = "Meilensteine"
                .cells(zeile, spalte + 2).interior.color = RGB(180, 180, 180)

                startMeilensteine = zeile


                For i = 1 To MilestoneDefinitions.Count
                    .cells(zeile, spalte).value = MilestoneDefinitions.getMilestoneDef(i).name
                    zeile = zeile + 1
                Next

                endZeile = zeile - 1

                If endZeile >= startMeilensteine Then
                    rng = CType(.range(.cells(startMeilensteine, spalte), .cells(endZeile, spalte)), Excel.Range)
                    appInstance.ActiveWorkbook.Names.Add(Name:="Meilensteine", RefersTo:=rng)

                End If

                If endZeile >= startZeile Then

                    If endZeile = startZeile Then
                        rng = CType(.cells(startZeile, spalte), Excel.Range)
                    Else
                        rng = CType(.range(.cells(startZeile, spalte), .cells(endZeile, spalte)), Excel.Range)
                    End If
                    appInstance.ActiveWorkbook.Names.Add(Name:="Phasen_Meilensteine", RefersTo:=rng)

                End If

                ' Schreiben des Delimiters
                .cells(zeile, spalte).value = delimiter
                zeile = zeile + 1


                ' ----------------------------------------- 
                ' Schreiben der Rollen und Kostenarten
                '

                .cells(zeile, spalte).value = "Rollen/Kostenarten"
                .cells(zeile, spalte).interior.color = RGB(180, 180, 180)

                zeile = zeile + 1
                startZeile = zeile
                startRollen = zeile

                For i = 1 To RoleDefinitions.Count
                    .cells(zeile, spalte).value = RoleDefinitions.getRoledef(i).name
                    zeile = zeile + 1
                Next

                endZeile = zeile - 1

                If endZeile >= startZeile Then

                    If endZeile = startZeile Then
                        rng = CType(.cells(startZeile, spalte), Excel.Range)
                    Else
                        rng = CType(.range(.cells(startZeile, spalte), .cells(endZeile, spalte)), Excel.Range)
                    End If
                    appInstance.ActiveWorkbook.Names.Add(Name:="Rollen", RefersTo:=rng)

                End If


                startKosten = zeile

                For i = 1 To CostDefinitions.Count - 1
                    ' die Personalkosten sind die letzte Kostenart, wird nicht mit aufgenommen, da sie 
                    ' automatisch berücksichtigt wird 
                    .cells(zeile, spalte).value = CostDefinitions.getCostdef(i).name
                    zeile = zeile + 1
                Next

                endZeile = zeile - 1
                If endZeile >= startKosten Then
                    rng = CType(.range(.cells(startKosten, spalte), .cells(endZeile, spalte)), Excel.Range)
                    appInstance.ActiveWorkbook.Names.Add(Name:="Kosten", RefersTo:=rng)

                End If

                If endZeile >= startZeile Then
                    rng = CType(.range(.cells(startZeile, spalte), .cells(endZeile, spalte)), Excel.Range)
                    appInstance.ActiveWorkbook.Names.Add(Name:="Rollen_Kostenarten", RefersTo:=rng)

                End If

                ' Schreiben des Delimiters
                .cells(zeile, spalte).value = delimiter
                zeile = zeile + 1


                ' -------------------------------------------
                ' Schreiben der Ampel-Farben 
                '


                .cells(zeile, spalte).value = "Ampel-Farben"
                .cells(zeile, spalte).interior.color = RGB(180, 180, 180)

                zeile = zeile + 1
                startZeile = zeile
                For i = 0 To 3
                    Select Case i
                        Case 0
                            .cells(zeile, spalte).value = "Ampel nicht bewertet"
                            .cells(zeile, spalte).interior.color = awinSettings.AmpelNichtBewertet
                        Case 1
                            .cells(zeile, spalte).value = "Ampel Grün"
                            .cells(zeile, spalte).interior.color = awinSettings.AmpelGruen
                        Case 2
                            .cells(zeile, spalte).value = "Ampel Gelb"
                            .cells(zeile, spalte).interior.color = awinSettings.AmpelGelb
                        Case 3
                            .cells(zeile, spalte).value = "Ampel Rot"
                            .cells(zeile, spalte).interior.color = awinSettings.AmpelRot
                    End Select
                    zeile = zeile + 1
                Next

                endZeile = zeile - 1

                If endZeile >= startZeile Then
                    rng = CType(.range(.cells(startZeile, spalte), .cells(endZeile, spalte)), Excel.Range)
                    appInstance.ActiveWorkbook.Names.Add(Name:="AmpelFarben", RefersTo:=rng)

                End If

                ' Schreiben des Delimiters
                .cells(zeile, spalte).value = delimiter
                zeile = zeile + 1


                ' ----------------------------------------- 
                ' Unsichtbarmachen des Tabellenblattes
                '

                .Visible = Excel.XlSheetVisibility.xlSheetHidden

            End With

        Catch ex As Exception
            appInstance.EnableEvents = formerEE
            Throw New ArgumentException("Fehler in awinExportProject, Schreiben Settings")
        End Try

        ' ----------------------------------------------
        ' jetzt werden die Termine weggeschrieben ....

        Try
            With appInstance.ActiveWorkbook.Worksheets("Termine")


                .Unprotect(Password:="x")       ' Blattschutz aufheben

                zeile = 1
                spalte = 1

            End With
        Catch ex As Exception
            With appInstance.ActiveWorkbook.Worksheets.Add
                .name = "Termine"
                ' Tabelle ErgebnTabelle muss hier eigentlich erzeugt werden
                appInstance.EnableEvents = formerEE
                Throw New ArgumentException("Fehler in awinExportProject, Schreiben Termine, Worksheet Termine existiert nicht")
            End With
        End Try

        ' --------------------------------------
        ' Worksheet Termine existriert jetzt ...

        With CType(appInstance.ActiveWorkbook.Worksheets("Termine"), Excel.Worksheet)

            .Unprotect(Password:="x")       ' Blattschutz aufheben

            Dim cphase As New clsPhase(hproj)
            Dim phaseName As String
            Dim r As Integer
            Dim cResult As New clsMeilenstein(parent:=cphase)
            Dim cBewertung As clsBewertung
            Dim phaseStart As Date
            Dim phaseEnde As Date
            Dim tbl As Excel.Range
            Dim itemNameID As String


            tbl = .Range("ErgebnTabelle")
            rowOffset = tbl.Row
            columnOffset = tbl.Column

            zeile = 0
          
            For p = 1 To hproj.CountPhases

                cphase = hproj.getPhase(p)

                ' Phasen-Name eintragen, Dauer einfärben
                itemNameID = cphase.nameID

                If awinSettings.zeitEinheit = "PM" Then
                    phaseStart = hproj.startDate.AddDays(cphase.startOffsetinDays)
                    phaseEnde = hproj.startDate.AddDays(cphase.startOffsetinDays + cphase.dauerInDays - 1)
                ElseIf awinSettings.zeitEinheit = "PW" Then
                    phaseStart = hproj.startDate.AddDays((cphase.relStart - 1) * 7)
                    phaseEnde = hproj.startDate.AddDays((cphase.relEnde - 1) * 7)
                ElseIf awinSettings.zeitEinheit = "PT" Then
                    phaseStart = hproj.startDate.AddDays(cphase.relStart - 1)
                    phaseEnde = hproj.startDate.AddDays(cphase.relEnde - 1)
                End If


                phaseName = cphase.name

                ' hier muss die Phase geschrieben werden


                If itemNameID = rootPhaseName Then

                    .Cells(rowOffset + zeile, columnOffset).value = elemNameOfElemID(rootPhaseName)

                Else
                    ' ur:06.05.2015: hier müssen die Einrückungen erfolgen

                    Dim indlevel As Integer = hproj.hierarchy.getIndentLevel(itemNameID)

                    '' ''Dim phstr As String = ""

                    ' '' '' in phstr werden nun soviele Leerzeichen hineingeschrieben, wie diese Phase Hierarchie-Stufen hat
                    '' ''For i = 1 To indlevel
                    '' ''    phstr = phstr & einrückJeStufe
                    '' ''Next

                    ' '' '' nun wird der PhasenName angehängt
                    '' ''phstr = phstr & elemNameOfElemID(itemNameID)

                    .Cells(rowOffset + zeile, columnOffset).value = elemNameOfElemID(itemNameID)
                    With CType(.Cells(rowOffset + zeile, columnOffset), Excel.Range)
                        .IndentLevel = indlevel * einrückTiefe
                        .HorizontalAlignment = Excel.XlHAlign.xlHAlignLeft
                    End With

                End If

                ' ur: 06.05 2015:Bezug fällt weg:.Cells(rowOffset + zeile, columnOffset + 2).value = ""
                .Cells(rowOffset + zeile, columnOffset + 2).value = phaseStart
                .Cells(rowOffset + zeile, columnOffset + 3).value = phaseEnde
                .Cells(rowOffset + zeile, columnOffset + 4).value = "0"
                .Cells(rowOffset + zeile, columnOffset + 4).interior.color = awinSettings.AmpelNichtBewertet
                .Cells(rowOffset + zeile, columnOffset + 5).value = " "
                .Cells(rowOffset + zeile, columnOffset + 6).value = " "

                zeile = zeile + 1

                For r = 1 To cphase.countMilestones
                    cResult = cphase.getMilestone(r)

                    cBewertung = cResult.getBewertung(1)
                    'Try
                    '    cBewertung = cResult.getBewertung(1)
                    'Catch ex As Exception
                    '    cBewertung = New clsBewertung
                    'End Try
                    ' --------------------------------------------------------------------------------
                    ' Termine müssen in Tabelle eingetragen werden
                    '----------------------------------------------------------------------------------

                    itemNameID = cResult.nameID

                    ' ur:06.05.2015: hier müssen die Einrückungen erfolgen

                    Dim indlevel As Integer = hproj.hierarchy.getIndentLevel(itemNameID)

                    .Cells(rowOffset + zeile, columnOffset).value = elemNameOfElemID(itemNameID)
                    With CType(.Cells(rowOffset + zeile, columnOffset), Excel.Range)

                        .IndentLevel = indlevel * einrückTiefe
                        .Font.Bold = True
                        .HorizontalAlignment = Excel.XlHAlign.xlHAlignLeft

                    End With
                    '.Cells(rowOffset + zeile, columnOffset + 2).value = cResult.getDate
                    .Cells(rowOffset + zeile, columnOffset + 3).value = cResult.getDate
                    .Cells(rowOffset + zeile, columnOffset + 4).value = cBewertung.colorIndex
                    .Cells(rowOffset + zeile, columnOffset + 4).interior.color = cBewertung.color
                    ' Zelle für Beschreibung in der Höhe anpassen, autom. Zeilenumbruch
                    .Cells(rowOffset + zeile, columnOffset + 5).value = cBewertung.description
                    .Cells(rowOffset + zeile, columnOffset + 5).WrapText = True

                    zeile = zeile + 1
                Next

            Next

            '' Blattschutz setzen
            '.Protect(Password:="x", UserInterfaceOnly:=True, DrawingObjects:=True, Contents:=True, Scenarios:=True)

        End With


        ' ----------------------------------------------
        ' jetzt werden die Attribute weggeschrieben ....

        Try
            With CType(appInstance.ActiveWorkbook.Worksheets("Attribute"), Excel.Worksheet)

                .Unprotect(Password:="x")       ' Blattschutz aufheben


                ' Projekt-Typ

                .Range("Projekt_Typ").Value = hproj.VorlagenName
                rng = .Range("Projekt_Typ")
                With rng
                    .HorizontalAlignment = Excel.XlHAlign.xlHAlignLeft
                    .IndentLevel = 1
                    .WrapText = False
                End With

                ' Status

                .Range("Status").Value = hproj.Status
                rng = .Range("Status")
                With rng
                    .HorizontalAlignment = Excel.XlHAlign.xlHAlignRight
                    .IndentLevel = 1
                    .WrapText = False
                End With

                ' Business_Unit

                .Range("Business_Unit").Value = hproj.businessUnit
                rng = .Range("Business_Unit")
                With rng
                    .HorizontalAlignment = Excel.XlHAlign.xlHAlignLeft
                    .IndentLevel = 1
                    .WrapText = False
                End With

                ' Strategischer Fit

                .Range("Strategischer_Fit").Value = hproj.StrategicFit
                rng = .Range("Strategischer_Fit")
                With rng
                    .HorizontalAlignment = Excel.XlHAlign.xlHAlignRight
                    .IndentLevel = 1
                    .WrapText = False
                End With

                ' Risiko

                .Range("Risiko").Value = hproj.Risiko
                rng = .Range("Risiko")
                With rng
                    .HorizontalAlignment = Excel.XlHAlign.xlHAlignRight
                    .IndentLevel = 1
                    .WrapText = False
                End With

                ' ur: 13.01.2015: Varianten_Name wird hier in das Tabellenblatt Attribute des Projekt-Steckbriefes eingetragen

                If Not IsNothing(hproj.variantName) And hproj.variantName <> "" Then

                    .Range("Variant_Name").Value = hproj.variantName
                    rng = .Range("Variant_Name")
                    With rng
                        .HorizontalAlignment = Excel.XlHAlign.xlHAlignLeft
                        .IndentLevel = 1
                        .WrapText = False
                    End With

                End If

                '' Blattschutz setzen
                '.Protect(Password:="x", UserInterfaceOnly:=True, DrawingObjects:=True, Contents:=True, Scenarios:=True)

            End With
        Catch ex As Exception
            '' Blattschutz setzen
            'appInstance.ActiveWorkbook.Worksheets("Attribute").Protect(Password:="x", UserInterfaceOnly:=True, DrawingObjects:=True, Contents:=True, Scenarios:=True)
            appInstance.EnableEvents = formerEE
            Throw New ArgumentException("Fehler in awinExportProject, Schreiben Attribute")
        End Try


        Try

            My.Computer.FileSystem.DeleteFile(fileName)
        Catch ex As Exception

        End Try

        Try

            appInstance.ActiveWorkbook.SaveAs(fileName, _
                                          ConflictResolution:=XlSaveConflictResolution.xlLocalSessionChanges
                                          )

        Catch ex As Exception
            appInstance.EnableEvents = formerEE
            Throw New ArgumentException("Fehler beim Datei-Schreiben")
        End Try


        appInstance.EnableEvents = formerEE


    End Sub
    '' '' '' ur: 13.01.2015: Funktion checkt ob ein String ein legaler DateiNamen ist( funktioniert aber nicht)
    ' '' ''Function IsLegalFileName(ByVal str As String) As Boolean
    ' '' ''    If (str Like "[/\:*?""<>]") Then
    ' '' ''        IsLegalFileName = True
    ' '' ''    Else
    ' '' ''        IsLegalFileName = False
    ' '' ''    End If
    ' '' ''End Function


    'ur: 13.01.2015: Funktion streicht die illegalen Zeigen heraus
    'entnommen von folgendem Link: http://www.jpsoftwaretech.com/excel-vba/validate-filenames/

    Function cleanFileName(stringToClean As String) As String
        ' remove illegal characters from filenames
        Dim newString As String

        newString = Replace(Replace(Replace(Replace(Replace(Replace(Replace(Replace(Replace(Replace(Replace(stringToClean, "|", ""), "[", ""), "]", ""), ">", ""), "<", ""), Chr(34), ""), "?", ""), "*", ""), ":", ""), "/", ""), "\", "")

        cleanFileName = newString

    End Function

    ' Vorbedingung: das Active-workbook ist bereits das ProjektDetail File 
    Public Sub awinStoreProjForEditRess(hproj As clsProjekt)
        Dim rng As Excel.Range
        Dim zeile As Integer, spalte As Integer
        Dim delimiter As String = "."
        Dim tmpStart As Date
        Dim pwd As String = "projekttafel"
        Dim laenge As Integer = hproj.anzahlRasterElemente
        Dim summenspalte As Integer
        Dim leadingColumns As Integer = 2

        Dim pstart As Integer = hproj.Start


        spalte = 1

        With CType(appInstance.Worksheets(arrWsNames(5)), Excel.Worksheet)
            ' Blattschutz aufheben 
            .Unprotect(Password:=pwd)

            ' Änderung 3.7.14 Bereich von Zelle(1,1) bis Zelle(2000,2000) schützen
            rng = .Range(.Cells(1, 1), .Cells(2000, 2000))
            rng.Clear()
            rng.Locked = True

            ' hier wird die Headerzeile beschrieben und sonstiges für Layout wichtiges Struktur 
            If laenge > 1 Then
                .Cells(1, 1).offset(0, leadingColumns).value = hproj.startDate
                .Cells(1, 1).offset(0, leadingColumns + 1).value = hproj.startDate.AddMonths(1)
            End If

            'For i = 1 To laenge
            '    '.cells(1, i + 2).value = StartofCalendar.AddMonths(pstart + i - 2).ToString("MMM yy")
            '    .Cells(1, i + 2).value = StartofCalendar.AddMonths(pstart + i - 2)
            'Next i

            rng = .Range(.Cells(1, 1).offset(0, leadingColumns), .Cells(1, 1).offset(0, leadingColumns + 1))

            If awinSettings.zeitEinheit = "PM" Then

                '.Cells(1, 1).value = "Monate"


                rng.NumberFormat = "mmm-yy"

                Dim destinationRange As Excel.Range

                destinationRange = .Range(.Cells(1, 1).offset(0, leadingColumns), _
                                          .Cells(1, 1).offset(0, leadingColumns + laenge - 1))
                summenspalte = spalte + 2 + laenge

                With destinationRange
                    .HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter
                    .VerticalAlignment = Excel.XlVAlign.xlVAlignBottom
                    .NumberFormat = "mmm-yy"
                    .WrapText = False
                    .Orientation = 90
                    .AddIndent = False
                    .IndentLevel = 0
                    .ShrinkToFit = False
                    .ReadingOrder = Excel.Constants.xlContext
                    .MergeCells = False
                    .Interior.Color = noshowtimezone_color
                End With

                rng.AutoFill(Destination:=destinationRange, Type:=Excel.XlAutoFillType.xlFillMonths)


            ElseIf awinSettings.zeitEinheit = "PW" Then
                .Cells(1, 1).value = "Wochen"
                For i = 1 To 210
                    CType(.Cells(1, i), Global.Microsoft.Office.Interop.Excel.Range).Value = StartofCalendar.AddDays((i - 1) * 7)
                Next

            ElseIf awinSettings.zeitEinheit = "PT" Then
                .Cells(1, 1).value = "Tage"
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
                Dim i As Integer, w As Integer
                i = 1
                For w = 1 To 30
                    For d = 0 To 4
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


            CType(.Rows(1), Global.Microsoft.Office.Interop.Excel.Range).RowHeight = awinSettings.zeilenhoehe1
            CType(.Range(.Cells(2, 1), .Cells(maxRows, maxColumns)), Global.Microsoft.Office.Interop.Excel.Range).RowHeight = awinSettings.zeilenhoehe2 * 0.5
            CType(.Range(.Cells(1, 3), .Cells(maxRows, maxColumns)), Global.Microsoft.Office.Interop.Excel.Range).ColumnWidth = awinSettings.spaltenbreite


            ' Ende einfügen aus Tabelle2 Activate Event, 3.7.2014

            ' hier wird die erste Zeile beschrieben 

            'For i = 1 To laenge
            '    '.cells(1, i + 2).value = StartofCalendar.AddMonths(pstart + i - 2).ToString("MMM yy")
            '    .Cells(1, i + 2).value = StartofCalendar.AddMonths(pstart + i - 2)
            'Next i

            'rng = .Range(.Cells(1, 3), .Cells(1, maxProjektdauer + 2))

            'Try
            '    rng.Columns.AutoFit()
            'Catch ex As Exception

            'End Try


            zeile = 2


            Dim k As Integer
            k = 0
            ' wenn es noch keine Phasen gibt: Projekt-Name eintragen, Dauer einfärben
            If hproj.CountPhases = 0 Then
                ' Projekt-Name eintragen, Dauer einfärben
                .Cells(zeile, spalte).value = hproj.name
                rng = .Range(.Cells(zeile, spalte + 2), .Cells(zeile, spalte + 1 + hproj.anzahlRasterElemente))
                rng.Interior.Color = hproj.farbe
                ' Kommentar am Anfang und am Ende der Phase mit Datum
                .Cells(zeile, spalte + 2).AddComment()
                With .Cells(zeile, spalte + 2).Comment
                    .Visible = False
                    .Text(Text:="Start:" & Chr(10) & hproj.startDate)
                    .Shape.ScaleHeight(0.45, Microsoft.Office.Core.MsoTriState.msoFalse)
                End With

                .Cells(zeile, spalte + 1 + hproj.anzahlRasterElemente).AddComment()
                With .Cells(zeile, spalte + 1 + hproj.anzahlRasterElemente).Comment
                    .Visible = False
                    .Text(Text:="Ende:" & Chr(10) & hproj.endeDate)
                    .Shape.ScaleHeight(0.45, Microsoft.Office.Core.MsoTriState.msoFalse)
                End With

                ' Änderung 3.7.14 Zellen sollen gesperrt sein für Änderungen
                rng.Locked = True

                zeile = zeile + 1
            End If

            ' hier über alle Phasen ... 
            Dim cphase As clsPhase
            Dim p As Integer
            Dim phasenFarbe As Object
            Dim values() As Double
            Dim ErgebnisListe As New Collection
            Dim anzahlItems As Integer
            Dim r As Integer
            Dim itemName As String
            Dim dimension As Integer
            'Dim atleastOne As Boolean = False


            For p = 1 To hproj.CountPhases
                cphase = hproj.getPhase(p)
                ' Phasen-Name eintragen, Dauer einfärben
                itemName = cphase.name

                Try
                    phasenFarbe = cphase.Farbe
                Catch ex As Exception
                    phasenFarbe = hproj.farbe
                End Try


                If itemName = elemNameOfElemID(rootPhaseName) Then
                    ' Projekt-Name eintragen, Dauer einfärben
                    .Cells(zeile, spalte).value = itemName
                    rng = .Range(.Cells(zeile, spalte + 2), .Cells(zeile, spalte + 1 + hproj.anzahlRasterElemente))
                    rng.Interior.Color = hproj.farbe

                    .Cells(zeile, spalte + 2).AddComment()
                    With .Cells(zeile, spalte + 2).Comment
                        .Visible = False
                        .Text(Text:="Start:" & Chr(10) & hproj.startDate)
                        .Shape.ScaleHeight(0.45, Microsoft.Office.Core.MsoTriState.msoFalse)
                    End With

                    .Cells(zeile, spalte + 1 + hproj.anzahlRasterElemente).AddComment()
                    With .Cells(zeile, spalte + 1 + hproj.anzahlRasterElemente).Comment
                        .Visible = False
                        .Text(Text:="Ende:" & Chr(10) & hproj.endeDate)
                        .Shape.ScaleHeight(0.45, Microsoft.Office.Core.MsoTriState.msoFalse)
                    End With

                    zeile = zeile + 1
                Else
                    ' Phasen Name eintragen, Dauer der Phase einfärben
                    .Cells(zeile, spalte).value = itemName
                    rng = .Range(.Cells(zeile, spalte + 1 + cphase.relStart), .Cells(zeile, spalte + 1 + cphase.relEnde))
                    rng.Interior.Color = phasenFarbe

                    .Cells(zeile, spalte + 1 + cphase.relStart).AddComment()
                    If cphase.relStart = cphase.relEnde Then
                        With .Cells(zeile, spalte + 1 + cphase.relStart).Comment
                            .Visible = False
                            .Text(Text:="Start:" & Chr(10) & cphase.getStartDate & Chr(10) & "Ende:" & Chr(10) & cphase.getEndDate)
                        End With
                    Else
                        With .Cells(zeile, spalte + 1 + cphase.relStart).Comment
                            .Visible = False
                            .Text(Text:="Start:" & Chr(10) & cphase.getStartDate)
                            .Shape.ScaleHeight(0.45, Microsoft.Office.Core.MsoTriState.msoFalse)
                        End With
                        .Cells(zeile, spalte + 1 + cphase.relEnde).AddComment()
                        With .Cells(zeile, spalte + 1 + cphase.relEnde).Comment
                            .Visible = False
                            .Text(Text:="Ende:" & Chr(10) & cphase.getEndDate)
                            .Shape.ScaleHeight(0.45, Microsoft.Office.Core.MsoTriState.msoFalse)
                        End With
                    End If

                    zeile = zeile + 1
                End If



                anzahlItems = cphase.countRoles


                Dim startEingabebereich As Integer = zeile
                Dim endeEingabebereich As Integer
                Dim eingabeBereich As Excel.Range

                For r = 1 To anzahlItems
                    itemName = cphase.getRole(r).name
                    dimension = cphase.getRole(r).getDimension
                    ReDim values(dimension)
                    values = cphase.getRole(r).Xwerte
                    .Cells(zeile, spalte + 1).value = itemName
                    rng = .Range(.Cells(zeile, spalte + 1 + cphase.relStart), .Cells(zeile, spalte + 1 + cphase.relStart + dimension))
                    rng.Value = values
                    zeile = zeile + 1
                Next r


                ' jetzt werden Kosten geschrieben 

                anzahlItems = cphase.countCosts

                For k = 1 To anzahlItems
                    itemName = cphase.getCost(k).name
                    dimension = cphase.getCost(k).getDimension
                    ReDim values(dimension)
                    values = cphase.getCost(k).Xwerte
                    .Cells(zeile, spalte + 1).value = itemName
                    rng = .Range(.Cells(zeile, spalte + 1 + cphase.relStart), .Cells(zeile, spalte + 1 + cphase.relStart + dimension))
                    rng.Value = values
                    zeile = zeile + 1
                Next k


                ' Änderung 3.7 diese Zeilen sollen änderbar sein 
                ' hier werden nun die Validation Kriterien für den Eingabebereich festgelegt 
                endeEingabebereich = zeile - 1

                If endeEingabebereich >= startEingabebereich Then
                    eingabeBereich = .Range(.Cells(startEingabebereich, spalte + 1 + cphase.relStart), _
                                        .Cells(endeEingabebereich, spalte + 1 + cphase.relStart + dimension))
                    eingabeBereich.Locked = False
                    eingabeBereich.Interior.Color = iProjektFarbe
                    Call InputZahlValidationforRange(eingabeBereich)
                End If

                zeile = zeile + 1
            Next p



            ' Phasen und Rollen sollen ja gar nicht mehr eingegeben werden können, deshalb ist das hier nicht mehr notwendig 

            ' hier werden nun die Phasen in einer Dropbox fixiert
            'rng = .Range(.Cells(2, 1), .Cells(200, 1))
            'InputValidationforRange(rng, 2, True)

            '' hier werden nun die Rollen und Kosten in Dropbox (2.Spalte) eingetragen
            'rng = .Range(.Cells(2, 2), .Cells(200, 2))
            'InputValidationforRange(rng, 1, True)


            ' Blattschutz aktivieren 
            .Protect(Password:=pwd, UserInterfaceOnly:=True, DrawingObjects:=True, Contents:=True, Scenarios:=True)

            ' jetzt noch die ersten beiden Spalten so dimensionieren, daß die Texte zu lesen sind
            CType(.Columns(1), Global.Microsoft.Office.Interop.Excel.Range).AutoFit()
            CType(.Columns(2), Global.Microsoft.Office.Interop.Excel.Range).AutoFit()

        End With



    End Sub
    '


    Public Sub awinReadProjFromEditRess(ByRef hproj As clsProjekt)
        Dim lastRow As Integer
        Dim rng As Excel.Range
        Dim zelle As Excel.Range
        Dim chkPhase As Boolean = True
        Dim Xwerte As Double()
        Dim crole As clsRolle
        Dim cphase As New clsPhase(hproj)
        Dim ccost As clsKostenart
        Dim phaseName As String
        Dim anfang As Integer, ende As Integer
        Dim farbeAktuell As Object
        Dim r As Integer, k As Integer
        'Dim zeile As Integer
        Dim newProj As New clsProjekt


        With appInstance.ActiveSheet

            'Dim valueRange As Excel.Range

            'zeile = 2
            'lastRow = .range(.Cells(1, 1), .cells(2000, 1)).End(XlDirection.xlUp).row
            lastRow = System.Math.Max(CType(.cells(2000, 1), Excel.Range).End(XlDirection.xlUp).Row, CType(.cells(2000, 2), Excel.Range).End(XlDirection.xlUp).Row) + 1
            rng = CType(.range(.cells(2, 1), .cells(lastRow, 1)), Excel.Range)
            'If .cells(zeile, 1).value <> hproj.name Then
            '    hproj.name = .cells(zeile, 1).value
            'End If

            'zeile = 3
            For Each zelle In rng
                Select Case chkPhase
                    Case True
                        ' hier wird die Phasen Information ausgelesen

                        If Len(CType(zelle.Value, String)) > 0 Then
                            phaseName = CType(zelle.Value, String).Trim

                            If Len(phaseName) > 0 Then

                                cphase = New clsPhase(hproj)

                                ' Auslesen der Phasen Dauer
                                anfang = 1
                                While zelle.Offset(0, anfang + 1).Interior.ColorIndex = -4142
                                    anfang = anfang + 1
                                End While

                                ende = anfang + 1
                                farbeAktuell = zelle.Offset(0, ende).Interior.Color
                                While zelle.Offset(0, ende + 1).Interior.Color = farbeAktuell
                                    ende = ende + 1
                                End While
                                ende = ende - 1

                                chkPhase = False


                                With cphase
                                    .nameID = hproj.hierarchy.findUniqueElemKey(phaseName, False)
                                    ' Änderung 28.11.13: jetzt wird die Phasen Länge exakt bestimmt , über startoffset in Tagen und dauerinDays als Länge
                                    Dim startOffset As Integer
                                    Dim dauerIndays As Integer
                                    startOffset = CInt(DateDiff(DateInterval.Day, hproj.startDate, hproj.startDate.AddMonths(anfang - 1)))
                                    dauerIndays = calcDauerIndays(hproj.startDate.AddDays(startOffset), ende - anfang + 1, True)


                                    .changeStartandDauer(startOffset, dauerIndays)
                                    .Offset = 0
                                End With

                            End If

                        End If



                    Case False ' auslesen Rollen- bzw. Kosten-Information

                        ' hier wird die Rollen bzw Kosten Information ausgelesen
                        Dim hname As String
                        Try
                            hname = CType(zelle.Offset(0, 1).Value, String).Trim
                        Catch ex1 As Exception
                            hname = ""
                        End Try


                        If Len(hname) > 0 Then

                            '
                            ' handelt es sich um die Ressourcen Definition?
                            '
                            If RoleDefinitions.Contains(hname) Then
                                Try
                                    r = CInt(RoleDefinitions.getRoledef(hname).UID)

                                    ReDim Xwerte(ende - anfang)


                                    'valueRange = .Range(zelle.Offset(0, anfang + 1), zelle.Offset(0, ende + 1))
                                    'Xwerte = CType(valueRange.Value, Double())

                                    For m = anfang To ende
                                        Xwerte(m - anfang) = CDbl(CType(zelle.Offset(0, m + 1), Excel.Range).Value)
                                    Next m

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

                            ElseIf CostDefinitions.Contains(hname) Then

                                Try

                                    k = CInt(CostDefinitions.getCostdef(hname).UID)

                                    ReDim Xwerte(ende - anfang)

                                    'valueRange = .Range(zelle.Offset(0, anfang + 1), zelle.Offset(0, ende + 1))
                                    'Xwerte = valueRange.Value

                                    For m = anfang To ende
                                        Xwerte(m - anfang) = CDbl(zelle.Offset(0, m + 1).Value)
                                    Next m

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


                        Else

                            chkPhase = True
                            hproj.AddPhase(cphase)

                        End If


                End Select
                'zeile = zeile + 1
            Next zelle



        End With



    End Sub


    ''' <summary>
    ''' ändert die Ressourcen Zuweisungen entsprechend den Eingaben in Editieren Ressourcen
    ''' Phase existiert und hat die gleiche Anzahl Monate: Behalten Start und Ende Datum
    ''' etwas anderes ist erst mal gar nicht erlaubt ; das wird schon durch die Protect Massnahme im Vorfeld erreicht 
    ''' </summary>
    ''' <param name="hproj"></param>
    ''' <remarks></remarks>
    Public Sub awinChangeProjFromEditRess(ByRef hproj As clsProjekt)
        Dim lastRow As Integer
        Dim rng As Excel.Range
        Dim zelle As Excel.Range
        Dim chkPhase As Boolean = True
        Dim Xwerte As Double()
        Dim crole As clsRolle
        Dim cphase As New clsPhase(hproj)
        Dim ccost As clsKostenart
        Dim phaseName As String, phaseNameID As String
        Dim anfang As Integer, ende As Integer
        Dim newProj As New clsProjekt
        Dim roleNr As Integer = 0, costNr As Integer = 0
        Dim checkliste As New SortedList(Of String, Integer)
        Dim lfdNr As Integer

        With appInstance.ActiveSheet


            lastRow = System.Math.Max(CType(.cells(2000, 1), Excel.Range).End(XlDirection.xlUp).Row, CType(.cells(2000, 2), Excel.Range).End(XlDirection.xlUp).Row) + 1
            rng = CType(.range(.cells(2, 1), .cells(lastRow, 1)), Excel.Range)

            For Each zelle In rng
                Select Case chkPhase
                    Case True
                        ' hier wird die Phasen Information ausgelesen

                        If Len(CType(zelle.Value, String)) > 0 Then
                            phaseName = CType(zelle.Value, String).Trim

                            ' prüfen, ob die schon mal da war 
                            If checkliste.ContainsKey(phaseName) Then
                                lfdNr = checkliste.Item(phaseName) + 1
                                checkliste.Item(phaseName) = lfdNr
                            Else
                                lfdNr = 1
                                checkliste.Add(phaseName, lfdNr)
                            End If

                            If Len(phaseName) > 0 Then

                                phaseNameID = calcHryElemKey(phaseName, False, lfdNr)
                                Try
                                    cphase = hproj.getPhaseByID(phaseNameID)
                                Catch ex As Exception

                                End Try

                                ' Auslesen der Phasen Dauer
                                anfang = cphase.relStart
                                ende = cphase.relEnde

                                chkPhase = False



                            End If

                        End If



                    Case False ' auslesen Rollen- bzw. Kosten-Information
                        ' durch EditStoreProjfor Ress wird sichergestellt, daß die 
                        ' rollen und Phasen in Ihrer Reihenfolge im Datenmodell ausgelesen werden 

                        ' hier wird die Rollen bzw Kosten Information ausgelesen
                        Dim hname As String
                        Try
                            hname = CType(zelle.Offset(0, 1).Value, String).Trim
                        Catch ex1 As Exception
                            hname = ""
                        End Try


                        If Len(hname) > 0 Then

                            '
                            ' handelt es sich um die Ressourcen Definition?
                            '
                            If RoleDefinitions.Contains(hname) Then

                                roleNr = roleNr + 1

                                Try
                                    crole = cphase.getRole(roleNr)

                                    ReDim Xwerte(ende - anfang)

                                    For m = anfang To ende

                                        Try
                                            Xwerte(m - anfang) = CDbl(zelle.Offset(0, m + 1).Value)
                                        Catch ex As Exception
                                            Xwerte(m - anfang) = 0.0
                                        End Try

                                    Next m

                                    With crole
                                        .Xwerte = Xwerte
                                    End With

                                Catch ex As Exception
                                    '
                                    ' handelt es sich um die Kostenart Definition?
                                    ' 


                                End Try

                            ElseIf CostDefinitions.Contains(hname) Then

                                costNr = costNr + 1

                                Try
                                    ccost = cphase.getCost(costNr)

                                    ReDim Xwerte(ende - anfang)

                                    For m = anfang To ende

                                        Try
                                            Xwerte(m - anfang) = CDbl(zelle.Offset(0, m + 1).Value)
                                        Catch ex As Exception
                                            Xwerte(m - anfang) = 0.0
                                        End Try

                                    Next m

                                    With ccost
                                        .Xwerte = Xwerte
                                    End With


                                Catch ex As Exception

                                End Try

                            End If


                        Else

                            chkPhase = True
                            roleNr = 0
                            costNr = 0

                        End If


                End Select

            Next zelle



        End With



    End Sub



    'Public Sub awinReadProjectTemplate(ByVal pname As String, ByVal intern As Boolean)


    '    Dim lastRow As Integer
    '    Dim rng As Excel.Range
    '    Dim zelle As Excel.Range
    '    Dim zeile As Integer, spalte As Integer
    '    Dim hproj As New clsProjektvorlage

    '    zeile = 1
    '    spalte = 1


    '    Try
    '        With appInstance.ActiveWorkbook.Worksheets("General Information")

    '            hproj.VorlagenName = CType(.cells(zeile, spalte + 1).value, String).Trim
    '            hproj.Schrift = .cells(zeile, spalte + 1).font.size
    '            hproj.Schriftfarbe = .cells(zeile, spalte + 1).font.color
    '            hproj.farbe = .cells(zeile, spalte + 1).interior.color

    '            ' earliest
    '            hproj.earliestStart = -6
    '            ' latest
    '            hproj.latestStart = 6


    '        End With
    '    Catch ex As Exception
    '        Throw New ArgumentException("Fehler beim auslesen General Information")
    '    End Try


    '    Try
    '        With appInstance.ActiveWorkbook.Worksheets("Project Needs")

    '            Dim chkPhase As Boolean = True
    '            Dim Xwerte As Double()
    '            Dim crole As clsRolle
    '            Dim cphase As New clsPhase(hproj, True)
    '            Dim ccost As clsKostenart
    '            Dim phaseName As String
    '            Dim anfang As Integer, ende As Integer
    '            Dim farbeAktuell As Object
    '            Dim r As Integer, k As Integer
    '            'Dim valueRange As Excel.Range

    '            zeile = 2

    '            lastRow = System.Math.Max(.cells(2000, 1).End(XlDirection.xlUp).row, .cells(2000, 2).End(XlDirection.xlUp).row) + 1
    '            rng = .range(.cells(2, 1), .cells(lastRow, 1))

    '            For Each zelle In rng
    '                Select Case chkPhase
    '                    Case True
    '                        ' hier wird die Phasen Information ausgelesen

    '                        If Len(CType(zelle.Value, String)) > 1 Then
    '                            phaseName = CType(zelle.Value, String).Trim

    '                            If Len(phaseName) > 0 Then

    '                                cphase = New clsPhase(hproj, True)

    '                                ' Auslesen der Phasen Dauer
    '                                anfang = 1
    '                                While zelle.Offset(0, anfang + 1).Interior.ColorIndex = -4142
    '                                    anfang = anfang + 1
    '                                End While

    '                                ende = anfang + 1
    '                                farbeAktuell = zelle.Offset(0, ende).Interior.Color
    '                                While zelle.Offset(0, ende + 1).Interior.Color = farbeAktuell
    '                                    ende = ende + 1
    '                                End While
    '                                ende = ende - 1

    '                                chkPhase = False


    '                                With cphase
    '                                    .name = phaseName
    '                                    ' Änderung 28.11.13: jetzt wird die Phasen Länge exakt bestimmt , über startoffset in Tagen und dauerinDays als Länge
    '                                    Dim startOffset As Integer = DateDiff(DateInterval.Day, StartofCalendar, StartofCalendar.AddMonths(anfang - 1))
    '                                    'Dim dauerIndays As Integer = DateDiff(DateInterval.Day, StartofCalendar.AddMonths(anfang - 1), _
    '                                    '                                                        StartofCalendar.AddMonths(ende).AddDays(-1)) + 1
    '                                    Dim dauerIndays As Integer = calcDauerIndays(StartofCalendar.AddDays(startOffset), ende - anfang + 1, True)
    '                                    .changeStartandDauer(startOffset, dauerIndays)

    '                                    .Offset = 0
    '                                End With

    '                            End If

    '                        End If



    '                    Case False ' auslesen Rollen- bzw. Kosten-Information

    '                        ' hier wird die Rollen bzw Kosten Information ausgelesen
    '                        Dim hname As String
    '                        Try
    '                            hname = CType(zelle.Offset(0, 1).Value, String).Trim
    '                        Catch ex1 As Exception
    '                            hname = ""
    '                        End Try


    '                        If Len(hname) > 0 Then

    '                            '
    '                            ' handelt es sich um die Ressourcen Definition?
    '                            '
    '                            If RoleDefinitions.Contains(hname) Then
    '                                Try
    '                                    r = RoleDefinitions.getRoledef(hname).UID

    '                                    ReDim Xwerte(ende - anfang)


    '                                    For m = anfang To ende
    '                                        Xwerte(m - anfang) = zelle.Offset(0, m + 1).Value
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
    '                                    '
    '                                    ' handelt es sich um die Kostenart Definition?
    '                                    ' 


    '                                End Try

    '                            ElseIf CostDefinitions.Contains(hname) Then

    '                                Try

    '                                    k = CostDefinitions.getCostdef(hname).UID

    '                                    ReDim Xwerte(ende - anfang)

    '                                    For m = anfang To ende
    '                                        Xwerte(m - anfang) = zelle.Offset(0, m + 1).Value
    '                                    Next m

    '                                    ccost = New clsKostenart(ende - anfang)
    '                                    With ccost
    '                                        .KostenTyp = k
    '                                        .Xwerte = Xwerte
    '                                    End With


    '                                    With cphase
    '                                        .AddCost(ccost)
    '                                    End With

    '                                Catch ex As Exception

    '                                End Try

    '                            End If


    '                        Else

    '                            chkPhase = True
    '                            hproj.AddPhase(cphase)

    '                        End If


    '                End Select
    '                zeile = zeile + 1
    '            Next zelle



    '        End With
    '    Catch ex As Exception
    '        Throw New ArgumentException("Fehler in awinImportProject, Lesen Project Needs")
    '    End Try


    '    ' hier werden die mit den Phasen verbundenen Results ausgelesen ...

    '    Try
    '        With appInstance.ActiveWorkbook.Worksheets("Settings")
    '            rng = .Range("Phasen")
    '            Dim rngZeile As Excel.Range
    '            Dim lastColumn As Integer
    '            Dim resultName As String = ""
    '            Dim phaseName As String
    '            Dim tmpPhase As New clsPhase(hproj, True)
    '            Dim tmpStr() As String
    '            Dim defaultOffset As Integer


    '            Dim anzTage As Integer

    '            For Each zelle In rng

    '                Try
    '                    phaseName = zelle.Value.trim

    '                    tmpPhase = hproj.getPhase(phaseName)
    '                    defaultOffset = tmpPhase.dauerInDays
    '                Catch ex As Exception

    '                End Try

    '                If Not tmpPhase Is Nothing Then

    '                    rngZeile = rng.Rows(zelle.Row)
    '                    lastColumn = .cells(zelle.Row, 2000).End(XlDirection.xlToLeft).column

    '                    Dim specified As Boolean
    '                    For i = 4 To lastColumn

    '                        specified = False
    '                        Try
    '                            resultName = .cells(zelle.Row, i).value.ToString.Trim

    '                            tmpStr = resultName.Split(New Char() {"(", ")"}, 10)

    '                            If tmpStr.Length > 1 Then

    '                                Try
    '                                    If awinSettings.offsetEinheit = "d" Then
    '                                        anzTage = CType(tmpStr(1), Integer)
    '                                    Else
    '                                        anzTage = CType(tmpStr(1), Integer) * 7
    '                                    End If

    '                                    resultName = tmpStr(0).Trim
    '                                    specified = True
    '                                Catch ex1 As Exception
    '                                    resultName = .cells(zelle.Row, i).value.ToString.Trim
    '                                    anzTage = defaultOffset
    '                                End Try

    '                            End If


    '                            Dim tmpResult As New clsResult(parent:=tmpPhase)

    '                            If resultName.Length > 0 Then
    '                                With tmpResult
    '                                    .name = resultName
    '                                    If specified Then
    '                                        .offset = anzTage
    '                                    Else
    '                                        .offset = defaultOffset
    '                                    End If
    '                                End With

    '                                tmpPhase.AddResult(tmpResult)

    '                            End If
    '                        Catch ex As Exception

    '                        End Try
    '                    Next

    '                End If

    '            Next

    '        End With
    '    Catch ex As Exception

    '    End Try



    '    Projektvorlagen.Add(hproj)


    'End Sub



    Public Function textZeitraum(start As Integer, ende As Integer) As String
        Dim htxt As String = " "
        Dim von As Date, bis As Date

        If start <= 0 Then
            start = 1
        End If

        Try
            With appInstance.Worksheets(arrWsNames(3))
                von = CDate(.cells(1, start).value)
                bis = CDate(.cells(1, ende).value)
                If start < ende Then
                    htxt = von.ToString("MMM yy") & " - " & bis.ToString("MMM yy")
                ElseIf start = ende Then
                    htxt = von.ToString("MMM yy")
                Else
                    htxt = bis.ToString("MMM yy") & " - " & von.ToString("MMM yy")
                End If
            End With
        Catch ex As Exception

        End Try


        textZeitraum = htxt

        'textZeitraum = StartofCalendar.AddMonths(start - 1).ToString("MMM yy") & " - " & _
        '                StartofCalendar.AddMonths(ende - 1).ToString("MMM yy")
    End Function

    ''' <summary>
    ''' gibt die Rasterspalte der Projekt-Tafel zurück, in der sich das angegebene Datum befindet 
    ''' die Einstellung awinsettings.zeiteinheit gibt dabei an, ob Monate , Wochen oder Tage das Raster sind
    ''' Aktuell werden nur Monate unterstützt
    ''' </summary>
    ''' <param name="datum"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function getColumnOfDate(ByVal datum As Date) As Integer
        Dim spalte As Integer = 1

        Select Case awinSettings.zeitEinheit
            Case "PM"
                spalte = CInt(DateDiff(DateInterval.Month, StartofCalendar, datum) + 1)
            Case "PW"
                Call MsgBox("noch nicht implementiert")
                spalte = 1
            Case "PT"
                Call MsgBox("noch nicht implementiert")
                spalte = 1
        End Select

        If spalte <= 0 Then
            getColumnOfDate = 1
        Else
            getColumnOfDate = spalte
        End If

    End Function

    Public Function getIndexBeauftragung(ByRef pHistorie As SortedList(Of Date, clsProjekt)) As Integer

        Dim abbruch As Boolean = False
        Dim tmpIndex As Integer = 0

        Dim anzSnapshots = pHistorie.Count

        ' jetzt wird der Planungs-Stand der Beauftragung gesucht 
        Do While pHistorie.ElementAt(tmpIndex).Value.Status <> ProjektStatus(1) And Not abbruch
            If tmpIndex + 1 < anzSnapshots Then
                tmpIndex = tmpIndex + 1
            Else
                abbruch = True
            End If
        Loop

        If abbruch Then
            ' es gibt keine Beauftragung ... 
            tmpIndex = -1
        Else
            ' index steht jetzt auf der Beauftragung 
        End If

        getIndexBeauftragung = tmpIndex

    End Function

    Public Function getIndexPrevFreigabe(ByRef pHistorie As SortedList(Of Date, clsProjekt), _
                                         ByVal currentIndex As Integer) As Integer

        Dim tmpIndex As Integer = currentIndex - 1
        Dim abbruch As Boolean = False

        If tmpIndex < 0 Then
            tmpIndex = -1
        Else

            Do While pHistorie.ElementAt(tmpIndex).Value.Status <> ProjektStatus(1) And Not abbruch
                If tmpIndex > 0 Then
                    tmpIndex = tmpIndex - 1
                Else
                    abbruch = True
                    tmpIndex = -1
                End If
            Loop

        End If

        getIndexPrevFreigabe = tmpIndex

    End Function

    Public Function getIndexNextFreigabe(ByRef pHistorie As SortedList(Of Date, clsProjekt), _
                                         ByVal currentIndex As Integer) As Integer

        Dim tmpIndex As Integer = currentIndex + 1
        Dim abbruch As Boolean = False

        If tmpIndex > pHistorie.Count - 1 Then
            tmpIndex = -1
        Else

            Do While pHistorie.ElementAt(tmpIndex).Value.Status <> ProjektStatus(1) And Not abbruch
                If tmpIndex < pHistorie.Count - 1 Then
                    tmpIndex = tmpIndex + 1
                Else
                    abbruch = True
                    tmpIndex = -1
                End If
            Loop

        End If

        getIndexNextFreigabe = tmpIndex

    End Function

    Public Function getIndexStartProject(ByRef pHistorie As SortedList(Of Date, clsProjekt)) As Integer

        ' noch nicht implementiert 
        getIndexStartProject = -1

    End Function

    Public Function getIndexEndProject(ByRef pHistorie As SortedList(Of Date, clsProjekt)) As Integer

        ' noch nicht implementiert 
        getIndexEndProject = -1

    End Function

    Public Function istLaufendesProjekt(ByRef hproj As clsProjekt) As Boolean

        Dim erg As Boolean = False

        Try
            With hproj
                If .Start <= getColumnOfDate(Date.Now) And
                    .Start + .anzahlRasterElemente - 1 >= getColumnOfDate(Date.Now) And _
                    .Status <> ProjektStatus(3) And _
                    .Status <> ProjektStatus(4) Then
                    erg = True
                End If
            End With
        Catch ex As Exception
            erg = False
        End Try

        istLaufendesProjekt = erg

    End Function

    ''' <summary>
    ''' errechnet die Kennung, die dem Chart als Namen mitgegeben wird; darf nicht größer als 31 in der Länge sein; 
    ''' das erlaubt Excel.Chart.name nicht
    ''' Typ ist entweder PF für Portfolio Kennung oder PR für die Projekt Charts  
    ''' 
    ''' </summary>
    ''' <param name="typ">ist Portfolio Chart (pf) oder Projekt-Chart (pr)</param>
    ''' <param name="index">gibt den Enumeration Wert an, der den Typ des Diagramms charakterisiert</param>
    ''' <param name="mycollection">enthält die Namen der Phasen/Rollen/Kostenarten bzw den Namen des Projektes oder 
    ''' wenn es sich um mehrere Projekte handelt: "x"</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function calcChartKennung(ByVal typ As String, ByVal index As Integer, ByVal mycollection As Collection) As String
        Dim IDkennung As String
        Dim cName As String = ""
        Dim breadcrumb As String = ""

        IDkennung = typ & "#" & index.ToString

        If typ = "pf" Then


            Try
                Select Case index
                    Case PTpfdk.Phasen

                        If mycollection.Count = PhaseDefinitions.Count Then
                            IDkennung = IDkennung & "#Alle"

                        Else

                            For i = 1 To mycollection.Count
                                cName = CStr(mycollection.Item(i)).Replace("#", "-")
                                ' der evtl vorhandenen Breadcrumb hat als Trennzeichen das #
                                Try
                                    IDkennung = IDkennung & "#" & cName
                                Catch ex As Exception
                                    IDkennung = IDkennung & "#"
                                End Try

                            Next

                        End If

                    Case PTpfdk.Meilenstein

                        For i = 1 To mycollection.Count
                            cName = CStr(mycollection.Item(i)).Replace("#", "-")
                            IDkennung = IDkennung & "#" & cName

                        Next

                    Case PTpfdk.Rollen

                        If mycollection.Count = RoleDefinitions.Count Then
                            IDkennung = IDkennung & "#Alle"

                        Else

                            For i = 1 To mycollection.Count
                                cName = CStr(mycollection.Item(i))
                                IDkennung = IDkennung & "#" & RoleDefinitions.getRoledef(cName).UID.ToString
                            Next

                        End If

                    Case PTpfdk.Kosten

                        If mycollection.Count = CostDefinitions.Count Then
                            IDkennung = IDkennung & "#Alle"

                        Else

                            For i = 1 To mycollection.Count
                                cName = CStr(mycollection.Item(i))
                                IDkennung = IDkennung & "#" & CostDefinitions.getCostdef(cName).UID.ToString
                            Next

                        End If

                    Case PTpfdk.ErgebnisWasserfall

                        If mycollection.Count > 0 Then
                            cName = CStr(mycollection.Item(1))
                            IDkennung = IDkennung & "#" & cName
                        End If

                    Case PTpfdk.Budget

                        If mycollection.Count > 0 Then
                            cName = CStr(mycollection.Item(1))
                            IDkennung = IDkennung & "#" & cName
                        End If

                End Select
            Catch ex As Exception

                IDkennung = IDkennung & "#?"
            End Try



        ElseIf typ = "pr" Then

            IDkennung = IDkennung & "#" & CStr(mycollection.Item(1))

        End If


        calcChartKennung = IDkennung


    End Function
    ''' <summary>
    ''' errechnet den für Showprojekte und AlleProjekte benötigten Schlüssel
    ''' setzt sich zusammen aus pName und variantName
    ''' </summary>
    ''' <param name="pName"></param>
    ''' <param name="variantName"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function calcProjektKey(ByVal pName As String, ByVal variantName As String) As String

        Dim trennzeichen As String = "#"

        ' Konsistenzbedingungen gewährleisten
        If IsNothing(pName) Then
            Throw New ArgumentException("Projekt-Name kann nicht Nothing sein")
        ElseIf pName.Length < 2 Then
            Throw New ArgumentException("Projekt-Name muss mindestens zwei Zeichen lang sein: " & pName)
        ElseIf IsNothing(variantName) Then
            variantName = ""
        End If

        calcProjektKey = pName & trennzeichen & variantName


    End Function

    ''' <summary>
    ''' errechnet den für Showprojekte und AlleProjekte benötigten Schlüssel
    ''' verwendet dazu die in hproj vorhandenen Attribute Name und variantName
    ''' </summary>
    ''' <param name="hproj"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function calcProjektKey(ByVal hproj As clsProjekt) As String

        Dim trennzeichen As String = "#"
        With hproj

            ' Konsistenzbedingungen gewährleisten
            If IsNothing(.name) Then
                Throw New ArgumentException("Projekt-Name kann nicht Nothing sein")
            ElseIf .name.Length < 2 Then
                Throw New ArgumentException("Projekt-Name muss mindestens zwei Zeichen lang sein: " & .name)
            ElseIf IsNothing(.variantName) Then
                .variantName = ""
            End If

            calcProjektKey = .name & trennzeichen & .variantName

        End With


    End Function

    ''' <summary>
    ''' errechnet den für Projekt-Varinte in der MongoDB verwendeten Schlüssel
    ''' ist aus historischen Gründen etwas anders als der Schlüssel in AlleProjekte 
    ''' </summary>
    ''' <param name="hproj"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function calcProjektKeyDB(ByVal hproj As clsProjekt) As String

        Dim tmpName As String
        With hproj


            ' Konsistenzbedingungen gewährleisten
            If IsNothing(.name) Then
                Throw New ArgumentException("Projekt-Name kann nicht Nothing sein")
            ElseIf .name.Length < 2 Then
                Throw New ArgumentException("Projekt-Name muss mindestens zwei Zeichen lang sein: " & .name)
            ElseIf IsNothing(.variantName) Then
                .variantName = ""
            End If

            If hproj.variantName <> "" And hproj.variantName.Trim.Length > 0 Then
                tmpName = calcProjektKey(hproj)
            Else
                tmpName = .name
            End If

            calcProjektKeyDB = tmpName

        End With


    End Function

    Public Function calcProjektKeyDB(ByVal pName As String, ByVal vName As String) As String

        Dim tmpName As String



        ' Konsistenzbedingungen gewährleisten
        If IsNothing(pName) Then
            Throw New ArgumentException("Projekt-Name kann nicht Nothing sein")
        ElseIf pName.Length < 2 Then
            Throw New ArgumentException("Projekt-Name muss mindestens zwei Zeichen lang sein: " & pName)
        ElseIf IsNothing(vName) Then
            vName = ""
        End If

        If vName <> "" And vName.Trim.Length > 0 Then
            tmpName = calcProjektKey(pName, vName)
        Else
            tmpName = pName
        End If

        calcProjektKeyDB = tmpName




    End Function

    ''' <summary>
    ''' gibt den Projekt-Namen zurück, der in dem Projekt-Key steckt 
    ''' </summary>
    ''' <param name="key"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function getPnameFromKey(ByVal key As String) As String
        Dim tmpStr(5) As String
        Dim trennzeichen As String = "#"

        tmpStr = key.Split(New Char() {CChar(trennzeichen)}, 4)
        getPnameFromKey = tmpStr(0)

    End Function

    ''' <summary>
    ''' gibt den Variant-Namen zurück, der in dem Projekt-Key steckt 
    ''' </summary>
    ''' <param name="key"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function getVariantnameFromKey(ByVal key As String) As String
        Dim tmpStr(5) As String
        Dim trennzeichen As String = "#"
        Dim tmpValue As String

        tmpStr = key.Split(New Char() {CChar(trennzeichen)}, 4)
        tmpValue = tmpStr(1)

        If IsNothing(tmpValue) Then
            tmpValue = ""
        End If

        getVariantnameFromKey = tmpValue

    End Function

    ''' <summary>
    ''' beschriftet ein Projekt mit seinen Phasen und / oder Meilenstein Namen
    ''' orientiert sich an dem Shape des Projektes , es wird nur beschriftet, was im Shape vorhanden ist
    ''' es kann angegeben werden, ob die Original- oder die Standard-Namen angezeigt werden sollen 
    ''' Alle Beschriftungen zusammen werden als ein zusammengesetztes Shape erzeugt
    ''' </summary>
    ''' <param name="projectShape"></param>
    ''' <param name="annotatePhases"></param>
    ''' <param name="annotateMilestones"></param>
    ''' <param name="showStdNames"></param>
    ''' <remarks></remarks>
    Public Sub annotateProject(ByVal projectShape As Excel.Shape, ByVal annotatePhases As Boolean, ByVal annotateMilestones As Boolean, _
                                   ByVal showStdNames As Boolean, ByVal showAbbrev As Boolean)


        Dim shapeSammlung As Excel.ShapeRange
        Dim descriptionGruppe As Excel.ShapeRange
        Dim descriptionShape As Excel.Shape
        Dim descriptionShapeName As String = "Description#" & projectShape.Name
        ' nimmt die Namen der Shapes auf, die zur Project Description gemacht werden sollen 
        Dim arrayOfNames() As String
        Dim anzahlElements As Integer

        Dim oldAlternativeText As String
        Dim oldTitle As String
        Dim oldName As String
        Dim hproj As clsProjekt
        Dim elemShape As Excel.Shape
        Dim txtShape As Excel.Shape
        Dim top As Single, left As Single, width As Single, height As Single
        Dim worksheetShapes As Excel.Shapes
        Dim nameID As String
        Dim description As String = ""
        
        Dim ok As Boolean
        Dim index As Integer = 0



        With CType(appInstance.Worksheets(arrWsNames(3)), Excel.Worksheet)
            worksheetShapes = .Shapes
        End With

        ' prüfen , ob es bereits ein Description Shape für dieses Projekt gibt; wenn ja, dann löschen
        Try
            descriptionShape = worksheetShapes.Item(descriptionShapeName)
            If Not IsNothing(descriptionShape) Then
                descriptionShape.Delete()
            End If
        Catch ex As Exception

        End Try

        With projectShape
            oldAlternativeText = .AlternativeText
            oldTitle = .Title
            oldName = .Name
        End With

        hproj = ShowProjekte.getProject(oldName)

        If isSingleProjectShape(projectShape) Then

            'Call MsgBox("es gibt keine Phasen oder Meilensteine zu beschriften ...")

        Else
            Try
                anzahlElements = projectShape.GroupItems.Count
                ReDim arrayOfNames(anzahlElements - 1)

                shapeSammlung = projectShape.Ungroup()

                ' hier muss dann die Aktion passieren
                index = 0
                For Each elemShape In shapeSammlung

                    ' zurücksetzen 
                    description = ""
                    txtShape = Nothing
                    ok = False
                    nameID = ""


                    If isPhaseType(kindOfShape(elemShape)) And annotatePhases Then
                        nameID = extractName(elemShape.Name, PTshty.phaseN)
                        ok = True

                    ElseIf isMilestoneType(kindOfShape(elemShape)) And annotateMilestones Then
                        nameID = extractName(elemShape.Name, PTshty.milestoneN)
                        ok = True

                    End If

                    If nameID = rootPhaseName Then
                        ok = False
                    End If

                    ' jetzt wird das Description Shape erzeugt 
                    If ok Then
                        ' nur, wenn es entweder ein Meilenstein oder eine Phase war ... 

                        description = hproj.hierarchy.getBestNameOfID(nameID, showStdNames, showAbbrev)

                        top = elemShape.Top
                        left = elemShape.Left
                        width = 30
                        height = 30

                        txtShape = worksheetShapes.AddLabel(MsoTextOrientation.msoTextOrientationHorizontal, _
                                                                left, top, width, height)

                        With txtShape
                            .TextFrame2.AutoSize = MsoAutoSize.msoAutoSizeShapeToFitText
                            .TextFrame2.WordWrap = MsoTriState.msoFalse
                            .TextFrame2.TextRange.Text = description
                            .TextFrame2.TextRange.Font.Size = hproj.Schrift - 2
                            .TextFrame2.MarginLeft = 0.1
                            .TextFrame2.MarginRight = 0.1
                            .TextFrame2.MarginTop = 0
                            .TextFrame2.MarginBottom = 0
                            .TextFrame2.VerticalAnchor = MsoVerticalAnchor.msoAnchorMiddle
                            .TextFrame2.HorizontalAnchor = MsoHorizontalAnchor.msoAnchorCenter

                            If description = "-" Then
                                .Fill.Visible = MsoTriState.msoFalse
                            Else
                                .Fill.Visible = MsoTriState.msoTrue
                                .Fill.ForeColor.RGB = RGB(255, 255, 255)
                                .Fill.Transparency = 0
                                .Fill.Solid()
                            End If


                        End With

                        ' jetzt muss das Shape noch in der Höhe richtig positioniert werden 
                        Dim diff As Single
                        Dim zeile As Integer

                        If istMeilensteinShape(elemShape) Then
                            'zeile = calcYCoordToZeile(elemShape.Top)
                            'txtShape.Top = CSng(calcZeileToYCoord(zeile)) + 1
                            txtShape.Top = elemShape.Top - txtShape.Height - 1
                            diff = (txtShape.Width - elemShape.Width) / 2
                            txtShape.Left = elemShape.Left - diff
                        Else
                            zeile = calcYCoordToZeile(elemShape.Top)
                            txtShape.Top = CSng(calcZeileToYCoord(zeile + 1)) - txtShape.Height - 1
                            'txtShape.Top = elemShape.Top + elemShape.Height
                            txtShape.Left = elemShape.Left
                        End If

                        ' jetzt wird das Shape aufgenommen 
                        arrayOfNames(index) = txtShape.Name
                        index = index + 1

                    End If

                Next

                Try
                    If index > 0 Then
                        If index < anzahlElements Then
                            anzahlElements = index
                            ReDim Preserve arrayOfNames(anzahlElements - 1)
                        End If

                        If anzahlElements > 1 Then
                            ' jetzt wird das neue zusammengesetzte Beschriftungs-Shape erzeugt ... 
                            descriptionGruppe = worksheetShapes.Range(arrayOfNames)
                            descriptionShape = descriptionGruppe.Group
                        Else
                            descriptionShape = worksheetShapes.Item(arrayOfNames(0))
                        End If

                        If Not IsNothing(descriptionShape) Then
                            With descriptionShape
                                .Name = descriptionShapeName
                                .AlternativeText = CInt(PTshty.beschriftung).ToString
                            End With
                        End If
                    End If
                Catch ex As Exception

                End Try

                ' hier muss das alte Shape wieder restauriert werden 
                projectShape = shapeSammlung.Group

                With projectShape
                    .Name = oldName
                    .AlternativeText = oldAlternativeText
                    .Title = oldTitle
                    hproj.shpUID = .ID.ToString
                End With

                ' jetzt muss das auch in der Liste Showprojekte eingetragen werden 
                ShowProjekte.AddShape(hproj.name, hproj.shpUID)

            Catch ex As Exception
                'Call MsgBox(ex.Message & vbLf & "... keine Phasen oder Meilensteine zu beschriften ...")
            End Try


        End If



    End Sub

    ''' <summary>
    ''' gibt true zurück, wenn es sich bei dem Element um einen Projekt-Meilenstein handelt 
    ''' false, sonst
    ''' </summary>
    ''' <param name="elemShape"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function istMeilensteinShape(ByVal elemShape As Excel.Shape) As Boolean

        With elemShape
            If .AlternativeText = CInt(PTshty.milestoneN).ToString Or _
                .AlternativeText = CInt(PTshty.milestoneE).ToString Then
                istMeilensteinShape = True
            Else
                istMeilensteinShape = False
            End If
        End With

    End Function

    ''' <summary>
    ''' gibt true zurück, wenn es sich bei dem Element um eine Projekt-Phase handelt
    ''' false, sonst
    ''' </summary>
    ''' <param name="elemShape"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function istPhasenShape(ByVal elemShape As Excel.Shape) As Boolean

        With elemShape
            If .AlternativeText = CInt(PTshty.phase1).ToString Or _
                    .AlternativeText = CInt(PTshty.phaseN).ToString Or _
                    .AlternativeText = CInt(PTshty.phaseE).ToString Then
                istPhasenShape = True
            Else
                istPhasenShape = False
            End If
        End With

    End Function

    ''' <summary>
    ''' aktualisiert bzw. zeigt das Status Fenster zur Ampel-Erläuterung eines Projektes 
    ''' </summary>
    ''' <param name="hproj">übergeben wird das betreffende Projekt</param>
    ''' <remarks></remarks>
    Public Sub updateStatusInformation(ByVal hproj As clsProjekt)

        Dim description As String

        Try

            description = hproj.ampelErlaeuterung


            If awinSettings.createIfNotThere Then
                ' prüfen, ob das Shape bereits angezeigt wird 
                Dim tstshpName = projectboardShapes.calcStatusShapeName(hproj.name, getColumnOfDate(hproj.timeStamp))
                If Not projectboardShapes.contains(tstshpName) Then
                    ' jetzt muss das Shape gezeichnet werden 
                    Dim tmpNr As Integer = 0

                    Dim formerEoU As Boolean = enableOnUpdate
                    Dim formerEE As Boolean = appInstance.EnableEvents

                    enableOnUpdate = False
                    appInstance.EnableEvents = False

                    Call zeichneStatusSymbolInPlantafel(hproj, tmpNr)


                    appInstance.EnableEvents = formerEE
                    enableOnUpdate = formerEoU

                End If
            End If

            With formStatus

                '.projectName.Text = hproj.name
                .projectName.Text = hproj.getShapeText
                .bewertungsText.Text = description

                If .Visible Then
                Else
                    .Visible = True
                    .Show()
                End If

            End With

        Catch ex As Exception


        End Try


    End Sub


    ''' <summary>
    ''' aktualisiert die Meilenstein Information; 
    ''' </summary>
    ''' <param name="hproj"></param>
    ''' <param name="milestoneNameID"></param>
    ''' <remarks></remarks>
    Public Sub updateMilestoneInformation(ByVal hproj As clsProjekt, ByVal milestoneNameID As String)


        Dim cMilestone As clsMeilenstein = Nothing
        Dim bewertung As New clsBewertung
        Dim ok As Boolean = True

        Dim projektName As String = ""
        Dim explanation As String = ""
        Dim milestoneName As String = "-"
        Dim breadCrumb As String = "-"
        Dim dateText As String = ""
        Dim farbe As System.Drawing.Color
        Dim lfdNr As Integer
        'Dim msNr As Integer, msNr2 As Integer

        Dim elemName As String
        Dim phaseName As String = ""
        Dim phaseName2 As String = ""
        Dim found As Boolean
        Dim tryoutName As String


        Try
            projektName = hproj.name
        Catch ex As Exception

        End Try

        Try
            cMilestone = hproj.getMilestoneByID(milestoneNameID)

            ' new Code

            ' es kann sein , dass es diesen Meilenstein nicht gibt, jedenfalls nicht mit der aktuellen lfdNr
            If IsNothing(cMilestone) Then
                lfdNr = lfdNrOfElemID(milestoneNameID)
                elemName = elemNameOfElemID(milestoneNameID)
                found = False
                Do While lfdNr > 1 And Not found
                    lfdNr = lfdNr - 1
                    tryoutName = calcHryElemKey(elemName, True, lfdNr)
                    cMilestone = hproj.getMilestoneByID(tryoutName)
                    If Not IsNothing(cMilestone) Then
                        found = True
                        milestoneNameID = tryoutName
                    End If
                Loop
            End If

            If Not IsNothing(cMilestone) Then
                ok = True

                If awinSettings.showOrigName Then
                    Dim tmpNode As clsHierarchyNode
                    tmpNode = hproj.hierarchy.nodeItem(milestoneNameID)
                    If Not IsNothing(tmpNode) Then
                        milestoneName = tmpNode.origName
                    Else
                        milestoneName = elemNameOfElemID(milestoneNameID)
                    End If
                Else
                    milestoneName = elemNameOfElemID(milestoneNameID)
                End If

                breadCrumb = hproj.hierarchy.getBreadCrumb(milestoneNameID).Replace("#", "-")
                lfdNr = lfdNrOfElemID(milestoneNameID)
                dateText = cMilestone.getDate.ToShortDateString

            Else
                ok = False
                milestoneName = "-"
                breadCrumb = "-"
                lfdNr = 0
                dateText = "-"
            End If

            ' Nicht ausliefern!
            ' das folgende dient auch Testzwecken ... wenn die beiden verschiedenen Arten, die Meilenstein Index Nummer zu bestimmen , 
            ' nicht übereinstimmen, wird einn  eine Fehlermeldung ausgegeben werden 

            'Try
            '    Dim cphase As clsPhase
            '    Dim parentPhaseNameID As String
            '    parentPhaseNameID = hproj.hierarchy.nodeItem(milestoneNameID).parentNodeKey
            '    cphase = hproj.getPhaseByID(parentPhaseNameID)
            '    msNr = cphase.getlfdNr(milestoneNameID)
            '    phaseName = cphase.name

            '    ' die elegantere, und neue Methode basierend auf der Hierarchie Liste
            '    msNr2 = hproj.hierarchy.nodeItem(milestoneNameID).indexOfElem
            '    phaseName2 = elemNameOfElemID(hproj.hierarchy.nodeItem(milestoneNameID).parentNodeKey)

            '    If msNr <> msNr2 Then
            '        Call MsgBox("hier stimmt was nicht: " & vbLf & msNr & " <> " & msNr2)
            '    End If

            '    If phaseName <> phaseName2 Then
            '        Call MsgBox("hier stimmt was nicht: " & vbLf & phaseName & " <> " & phaseName2)
            '    End If
            'Catch ex1 As Exception
            '    Call MsgBox("hier stimmt was nicht: (mit Fehler)" & vbLf & msNr & " <> " & msNr2 & vbLf & _
            '                phaseName & " <> " & phaseName2)
            'End Try




        Catch ex As Exception
            explanation = milestoneNameID & " existiert nicht"
            ok = False
        End Try


        If ok Then

            If awinSettings.createIfNotThere Then
                ' prüfen, ob das Shape bereits angezeigt wird 
                Dim tstshpName = projectboardShapes.calcMilestoneShapeName(projektName, milestoneNameID)
                If Not projectboardShapes.contains(tstshpName) Then
                    ' jetzt muss das Shape gezeichnet werden 
                    Dim tmpNr As Integer = 1
                    Dim tmpCollection As New Collection
                    tmpCollection.Add(milestoneNameID, milestoneNameID)

                    Dim formerEoU As Boolean = enableOnUpdate
                    Dim formerEE As Boolean = appInstance.EnableEvents
                    enableOnUpdate = False
                    appInstance.EnableEvents = False

                    Call zeichneMilestonesInProjekt(hproj, tmpCollection, 4, 0, 0, False, tmpNr, False)

                    appInstance.EnableEvents = formerEE
                    enableOnUpdate = formerEoU

                End If
            End If

            If cMilestone.bewertungsListe.Count > 0 Then

                Dim hb As clsBewertung = cMilestone.bewertungsListe.ElementAt(0).Value
                farbe = System.Drawing.Color.FromArgb(CInt(hb.color))

                explanation = hb.description


            Else

                farbe = System.Drawing.Color.FromArgb(CInt(awinSettings.AmpelNichtBewertet))
                explanation = "es existiert noch keine Bewertung ...."

            End If

            dateText = cMilestone.getDate.ToShortDateString


        End If


        With formMilestone

            .milestoneNameID = milestoneNameID
            .curProject = hproj

            .projectName.Text = hproj.getShapeText
            .breadCrumb.Text = breadCrumb

            .resultDate.Text = dateText
            .resultName.Text = milestoneName

            'If lfdNr > 1 Then
            '    .lfdNr.Text = lfdNr.ToString("0#")
            'Else
            '    .lfdNr.Text = ""
            'End If

            .bewertungsText.Text = explanation

            If .Visible Then
            Else
                .Visible = True
                .Show()
            End If

        End With


    End Sub

    ''' <summary>
    ''' überprüft, ob sich das Shape in Breite bzw. Position verändert hat
    ''' wenn ja, werden die Phasen Daten entsprechend geändert 
    ''' </summary>
    ''' <param name="phaseShape"></param>
    ''' <remarks></remarks>
    Public Sub updatePhaseStartDuration(ByVal phaseShape As Excel.Shape)

        Dim tmpstr() As String
        Dim projectName As String
        Dim phaseNameID As String
        Dim phaseNr As Integer
        Dim cPhase As clsPhase
        Dim zeilenoffset As Integer
        Dim oldType As String

        Dim ok As Boolean = True
        Dim hproj As New clsProjekt

        'Dim top1 As Double, top2 As Double, left1 As Double, left2 As Double
        'Dim ld As Double

        Dim phasenStart As Integer
        Dim phasenDauer As Integer
        Dim actionNeeded As Boolean = False

        Dim sollLeft As Double, sollWidth As Double, sollTop As Double, sollHeight As Double
        Dim istLeft As Double, istWidth As Double

        Dim projectShapes As Excel.Shapes = CType(appInstance.Worksheets(arrWsNames(3)), Excel.Worksheet).Shapes
        Dim suchstring = phaseShape.Name

        Dim projectShape As Excel.ShapeRange, shapeSammlung As Excel.ShapeRange
        Dim pShape As Excel.Shape


        With phaseShape

            istLeft = .Left
            istWidth = .Width

            tmpstr = .Name.Split(New Char() {CChar("#")}, 10)
            projectName = tmpstr(0)
            oldType = .AlternativeText

        End With

        Try
            hproj = ShowProjekte.getProject(projectName)
        Catch ex As Exception
            hproj = Nothing
            ok = False
        End Try




        If ok Then



            cPhase = New clsPhase(hproj)
            Try
                phaseNameID = tmpstr(1).Trim
                cPhase = hproj.getPhaseByID(phaseNameID)

                ' jetzt muss geprüft werden, ob sich die Start-Position, Länge oder Endposition geändert hat 

                If awinSettings.drawphases = True Then

                    projectShape = projectShapes.Range(projectName)
                    shapeSammlung = projectShape.Ungroup()

                    ' muss hier gemacht werden, weil die Koordinaten andere sind, wenn das Shape noch gruppiert ist 
                    With phaseShape

                        istLeft = .Left
                        istWidth = .Width

                    End With


                    phaseNr = CInt(tmpstr(2))
                    zeilenoffset = 0
                    Call hproj.CalculateShapeCoord(phaseNr, zeilenoffset, sollTop, sollLeft, sollWidth, sollHeight)

                    'Dim korrfaktorleft As Double = istLeft - sollLeft
                    'Dim korrfakttorwidth As Double = istWidth - sollWidth
                    'Dim korr1 As Double = istLeft / sollLeft
                    'Dim korr2 As Double = istWidth / sollWidth
                    '' ggf Height und Top anpassen 

                    'sollLeft = istLeft
                    'sollWidth = istWidth
                    With phaseShape

                        If .Height <> sollHeight Then
                            .Height = CSng(sollHeight)
                        End If

                        If .Top <> sollTop Then
                            .Top = CSng(sollTop)
                        End If

                    End With

                    ' jetzt wird wieder regruppiert
                    pShape = shapeSammlung.Regroup
                    With pShape

                        .Name = projectName
                        .AlternativeText = oldType
                        hproj.shpUID = .ID.ToString

                    End With

                    ' jetzt muss das auch in der Liste Showprojekte eingetragen werden 
                    ShowProjekte.AddShape(projectName, hproj.shpUID)


                Else
                    'Call cPhase.calculateLineCoord(hproj.tfZeile, 1, 1, top1, left1, top2, left2, ld)
                    Call cPhase.calculatePhaseShapeCoord(sollTop, sollLeft, sollWidth, sollHeight)
                    'sollLeft = left1
                    'sollWidth = left2 - left1
                    'sollTop = top1 - boxHeight * 0.3 * 0.5
                End If

                If System.Math.Abs(sollLeft - istLeft) > 0.5 Or _
                    System.Math.Abs(sollWidth - istWidth) > 0.5 Then
                    ' dann muss der Start bzw. die Duration geändert werden  

                    phasenStart = CInt(365 * istLeft / (12 * boxWidth)) - CInt(DateDiff(DateInterval.Day, StartofCalendar, hproj.startDate))
                    If phasenStart < 0 Then
                        phasenStart = 0
                    End If

                    phasenDauer = CInt(365 * (istLeft + istWidth) / (12 * boxWidth)) - CInt(DateDiff(DateInterval.Day, StartofCalendar, hproj.startDate)) - phasenStart

                    Call cPhase.changeStartandDauer(phasenStart, phasenDauer)
                    actionNeeded = True
                    ' dann der Ordnung halber auf die Soll-Werte setzen 

                    With phaseShape
                        .Top = CSng(sollTop)
                    End With



                Else
                    ' dann der Ordnung halber auf die Soll-Werte setzen 

                    If awinSettings.drawphases = False Then

                        With phaseShape

                            .Left = CSng(sollLeft)
                            .Width = CSng(sollWidth)
                            .Top = CSng(sollTop)

                        End With

                    End If

                End If


            Catch ex As Exception
                phaseNameID = ""
                ok = False
            End Try


        Else
            'Call MsgBox("keine Information abrufbar ...")
        End If

        If actionNeeded Then
            Call awinNeuZeichnenDiagramme(1)
        End If




    End Sub


    'Public Sub updatePhaseInformation(ByVal phaseShape As Excel.Shape)

    '    Dim tmpstr() As String
    '    Dim projectName As String
    '    Dim phaseName As String
    '    Dim cPhase As clsPhase

    '    Dim ok As Boolean = True
    '    Dim hproj As New clsProjekt

    '    Dim phaseStartdate As Date
    '    Dim phaseEnddate As Date
    '    Dim phaseDauerDays As Integer



    '    With phaseShape

    '        tmpstr = .Name.Split(New Char() {CChar("#")}, 10)
    '        projectName = tmpstr(0)

    '    End With

    '    Try
    '        hproj = ShowProjekte.getProject(projectName)
    '    Catch ex As Exception
    '        hproj = Nothing
    '        ok = False
    '    End Try

    '    If ok Then

    '        cPhase = New clsPhase(hproj)
    '        Try
    '            phaseName = tmpstr(1).Trim
    '            cPhase = hproj.getPhase(phaseName)
    '            'phaseStartdate = hproj.startDate.AddMonths(cPhase.relStart - 1)

    '            phaseStartdate = cPhase.getStartDate
    '            phaseEnddate = cPhase.getEndDate
    '            phaseDauerDays = cPhase.dauerInDays


    '            With formPhase

    '                If specialListofPhases.Contains(phaseName) Then

    '                    .projectName.Text = projectName
    '                    .phaseName.Text = phaseName
    '                    .Height = 440
    '                    .lessonsLearnedControl.Visible = True
    '                    .erlaeuterung.Visible = True
    '                    .erlaeuterung.Text = " ... hier werden die Prämissen angezeigt bzw. verändert "
    '                    .explSonderabl.Text = "Sonderabläufe der Phase " & phaseName & _
    '                        ", Projekt " & projectName
    '                    .explEnabler.Text = "Enabler der Phase " & phaseName & _
    '                        ", Projekt " & projectName
    '                    .explRisiken.Text = "Zusatzrisiken der Phase " & phaseName & _
    '                        ", Projekt " & projectName

    '                    .phaseStart.Text = phaseStartdate.ToShortDateString
    '                    .phaseStart.TextAlign = HorizontalAlignment.Left

    '                    .phaseEnde.Text = phaseEnddate.ToShortDateString
    '                    .phaseEnde.TextAlign = HorizontalAlignment.Right

    '                    .phaseDauer.Text = phaseDauerDays.ToString & " Tage"
    '                    .phaseDauer.TextAlign = HorizontalAlignment.Center


    '                    If .Visible Then
    '                    Else
    '                        .Visible = True
    '                        .Show()
    '                    End If


    '                Else

    '                    .projectName.Text = projectName
    '                    .phaseName.Text = phaseName
    '                    .Height = 220
    '                    .erlaeuterung.Visible = False

    '                    .phaseStart.Text = phaseStartdate.ToShortDateString
    '                    .phaseStart.TextAlign = HorizontalAlignment.Left

    '                    .phaseEnde.Text = phaseEnddate.ToShortDateString
    '                    .phaseEnde.TextAlign = HorizontalAlignment.Right

    '                    .phaseDauer.Text = phaseDauerDays.ToString & " Tage"
    '                    .phaseDauer.TextAlign = HorizontalAlignment.Center

    '                    If .Visible Then
    '                    Else
    '                        .Visible = True
    '                        .Show()
    '                    End If


    '                End If


    '            End With


    '        Catch ex As Exception
    '            phaseName = ""
    '            ok = False
    '        End Try


    '    Else
    '        'Call MsgBox("keine Information abrufbar ...")
    '    End If






    'End Sub

    ''' <summary>
    ''' aktualisiert die Phasen-Information zu dem Projekt; 
    ''' wenn die Phase nicht existiert, werden  
    ''' </summary>
    ''' <param name="hproj"></param>
    ''' <param name="phaseNameID"></param>
    ''' <remarks></remarks>
    Public Sub updatePhaseInformation(ByVal hproj As clsProjekt, phaseNameID As String)

        Dim projectName As String = hproj.name
        Dim lfdNr As Integer
        Dim elemName As String
        Dim tryoutName As String
        Dim found As Boolean
        Dim cPhase As clsPhase

        Dim ok As Boolean = True

        Dim phaseName As String = "-"
        Dim breadCrumb As String = "-"
        Dim startdateText As String = "-"
        Dim enddateText As String = "-"
        Dim dauerText = "-"


        Try
            cPhase = hproj.getPhaseByID(phaseNameID)
            ' es kann sein , dass es diese Phase nicht gibt, jedenfalls nicht mit der aktuellen lfdNr
            If IsNothing(cPhase) Then
                lfdNr = lfdNrOfElemID(phaseNameID)
                elemName = elemNameOfElemID(phaseNameID)
                found = False
                Do While lfdNr > 1 And Not found
                    lfdNr = lfdNr - 1
                    tryoutName = calcHryElemKey(elemName, False, lfdNr)
                    cPhase = hproj.getPhaseByID(tryoutName)
                    If Not IsNothing(cPhase) Then
                        found = True
                        phaseNameID = tryoutName
                    End If
                Loop
            End If

            If Not IsNothing(cPhase) Then
                ok = True
                If awinSettings.showOrigName Then
                    Dim tmpNode As clsHierarchyNode
                    tmpNode = hproj.hierarchy.nodeItem(cPhase.nameID)
                    If Not IsNothing(tmpNode) Then
                        phaseName = tmpNode.origName
                    Else
                        phaseName = elemNameOfElemID(phaseNameID)
                    End If
                Else
                    phaseName = elemNameOfElemID(phaseNameID)
                End If

                breadCrumb = hproj.hierarchy.getBreadCrumb(phaseNameID).Replace("#", "-")
                lfdNr = lfdNrOfElemID(phaseNameID)
                startdateText = cPhase.getStartDate.ToShortDateString
                enddateText = cPhase.getEndDate.ToShortDateString
                dauerText = cPhase.dauerInDays.ToString & " Tage"

            Else
                ok = False
                phaseName = "-"
                breadCrumb = "-"
                lfdNr = 0
                startdateText = "-"
                enddateText = "-"
                dauerText = "-"
            End If



        Catch ex As Exception
            phaseName = ""
            ok = False
        End Try


        If ok Then
            If awinSettings.createIfNotThere And Not awinSettings.drawphases Then
                ' prüfen, ob das Shape bereits angezeigt wird 
                Dim tstshpName = projectboardShapes.calcPhaseShapeName(projectName, phaseNameID)
                If Not projectboardShapes.contains(tstshpName) Then
                    ' jetzt muss das Shape gezeichnet werden 
                    Dim tmpNr As Integer = 1
                    Dim tmpCollection As New Collection
                    tmpCollection.Add(phaseNameID, phaseNameID)

                    Dim formerEoU As Boolean = enableOnUpdate
                    Dim formerEE As Boolean = appInstance.EnableEvents

                    enableOnUpdate = False
                    appInstance.EnableEvents = False

                    Call zeichnePhasenInProjekt(hproj, tmpCollection, False, tmpNr)

                    appInstance.EnableEvents = formerEE
                    enableOnUpdate = formerEoU

                End If
            End If
        End If




        With formPhase

            .phaseNameID = phaseNameID
            .curProject = hproj

            .projectName.Text = hproj.getShapeText
            .breadCrumb.Text = breadCrumb

            .phaseName.Text = phaseName


            .phaseStart.Text = startdateText
            .phaseStart.TextAlign = HorizontalAlignment.Left

            .phaseEnde.Text = enddateText
            .phaseEnde.TextAlign = HorizontalAlignment.Right

            .phaseDauer.Text = dauerText
            .phaseDauer.TextAlign = HorizontalAlignment.Center

            If .Visible Then
            Else
                .Visible = True
                .Show()
            End If



        End With


    End Sub

    ''' <summary>
    ''' bringt eine Liste von Phasen Namen zurück, die in den beiden Projekten einander identisch sind 
    ''' Wenn die Collection leer ist, dann unterscheiden sich beide Projekte in allen Phasen 
    ''' </summary>
    ''' <param name="hproj">Projekt 1</param>
    ''' <param name="cproj">Projekt 2</param>
    ''' <returns>Liste von Phasen Namen, die identisch sind </returns>
    ''' <remarks></remarks>
    Public Function getPhasenUnterschiede(ByVal hproj As clsProjekt, ByVal cproj As clsProjekt) As Collection
        Dim noColorCollection As New Collection
        Dim hphase As clsPhase, cphase As clsPhase
        Dim phaseNameID As String


        For p = 1 To hproj.CountPhases

            Try
                If p = 1 Then
                    hphase = hproj.getPhase(1)
                    cphase = cproj.getPhase(1)

                    If hphase.startOffsetinDays = cphase.startOffsetinDays And _
                            hphase.dauerInDays = cphase.dauerInDays Then
                        Try
                            ' in diesem Fall müssen beide Phase(1) Namen, die ja evtl unterschiedlich sind, aufgenommen werden 
                            ' 'nderung 13.4.15 jetzt muss das nur noch einmal aufgenommen werden ..., da die Root Phase jetzt immer gleich heisst
                            'noColorCollection.Add(hphase.name, hphase.name)
                            noColorCollection.Add(cphase.nameID, cphase.nameID)
                        Catch ex As Exception

                        End Try
                    End If
                Else

                    hphase = hproj.getPhase(p)
                    phaseNameID = hphase.nameID

                    Try
                        cphase = cproj.getPhaseByID(phaseNameID)

                        If hphase.startOffsetinDays = cphase.startOffsetinDays And _
                            hphase.dauerInDays = cphase.dauerInDays Then
                            Try
                                noColorCollection.Add(phaseNameID, phaseNameID)
                            Catch ex As Exception

                            End Try
                        End If
                    Catch ex As Exception
                        ' in diesem Fall gibt es die Phase in hproj, nicht aber in cproj ... 
                        ' das heisst, es muss farbig gezeichnet werden ... also nicht in NoColorCollection aufnehmen 
                    End Try
                End If

            Catch ex As Exception
                ' in diesem Fall ist gar nichts zu tun ... 
            End Try


        Next

        getPhasenUnterschiede = noColorCollection


    End Function

    ''' <summary>
    ''' speichert die aktuelle Konstellation in currentProjektListe in eine Konstellation
    ''' wenn die ImportProjekte vom Typ clsProjekteAlle übergeben wird, dann wird die hergenommen, um die Constellation aufzubauen 
    ''' </summary>
    ''' <param name="constellationName"></param>
    ''' <remarks></remarks>
    Public Sub storeSessionConstellation(ByRef currentProjektListe As clsProjekte, ByVal constellationName As String, _
                                         Optional ByVal ImportProjekte As clsProjekteAlle = Nothing)

        'Dim request As New Request(awinSettings.databaseName)


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

        If Not IsNothing(ImportProjekte) Then
            For Each kvp As KeyValuePair(Of String, clsProjekt) In ImportProjekte.liste
                newConstellationItem = New clsConstellationItem
                With newConstellationItem
                    .projectName = kvp.Value.name
                    .show = True
                    .Start = kvp.Value.startDate
                    .variantName = kvp.Value.variantName
                    .zeile = kvp.Value.tfZeile
                End With
                newC.Add(newConstellationItem)
            Next
        Else
            For Each kvp As KeyValuePair(Of String, clsProjekt) In currentProjektListe.Liste
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
        End If



        Try
            projectConstellations.Add(newC)

        Catch ex As Exception
            Call MsgBox("Fehler bei Add projectConstellations in awinStoreConstellations")
        End Try




    End Sub

    ''' <summary>
    ''' aktiviert die angegebene Projekt-Variante und zeichnet das entsprechende Shape in der Projekt-Tafel 
    ''' selektiert ggf das Shape, um die Aktualisierung gleich durchzuführen 
    ''' wenn replaceAnyhow, dann wird auf alle Fälle ersetzt, andernfalls nur, wenn der Varianten-Name nicht schon geladen ist
    ''' </summary>
    ''' <param name="pname"></param>
    ''' <param name="newVariant"></param>
    ''' <param name="selectIT" >gibt an, ob das Shape gleich selektiert werden soll</param>
    ''' <param name="replaceAnyhow">gibt an, ob die Ersetzung auf alle Fälle erfolgen soll oder nur wenn nicht diese Variante 
    ''' bereits in Showprojekte geladen ist </param>
    ''' <param name="tfzeile">gibt an, ab welcher Zeile auf der Projekttafel versucht werden soll, zu zeichnen
    ''' </param>
    ''' <remarks></remarks>
    Sub replaceProjectVariant(ByVal pname As String, ByVal newVariant As String, _
                              ByVal selectIT As Boolean, ByVal replaceAnyhow As Boolean, _
                              ByVal tfzeile As Integer)

        Dim newProj As clsProjekt
        Dim hproj As clsProjekt
        Dim key As String = calcProjektKey(pname, newVariant)
        'Dim tfzeile As Integer = 0
        'Dim projectshape As Excel.ShapeRange

        Dim phaseList As New Collection
        Dim milestoneList As New Collection


        ' gibt es die neue Variante überhaupt ? 
        If AlleProjekte.Containskey(key) Then
            newProj = AlleProjekte.getProject(key)

            ' jetzt muss die bisherige Variante aus Showprojekte rausgenommen werden ..
            If ShowProjekte.contains(pname) Then
                hproj = ShowProjekte.getProject(pname)

                ' welche Phasen werden angezeigt , welche Meilensteine werden angezeigt ? 
                phaseList = projectboardShapes.getPhaseList(pname)
                milestoneList = projectboardShapes.getMilestoneList(pname)

                ' prüfen, ob es überhaupt eine andere Variante ist 
                ' Änderung 09.10.14: das sollte dann ein Abbruch-Kriterium sein, wenn nicht ohnehin ersetzt werden soll 
                ' denn wenn das Projekt aus der Datenbank neu geladen wird, kann es ggf unterschiedlich sein; 
                ' also sollte es bei replaceAnyhow auf alle Fälle geladen werden 
                If hproj.variantName = newVariant And Not replaceAnyhow Then
                    Exit Sub
                End If

                ' bestimme die bisher angezeigten Phasen und Meilensteine 


                tfzeile = hproj.tfZeile

                ' die Darstellung in der Projekt-Tafel löschen
                Call clearProjektinPlantafel(pname)

                ' Änderung tk 4.7.15 erst clear auf Tafel, dann Remove aus Showprojekte 
                ' andernfalls macht der Clear mit Röntgen-Blick Schwierigkeiten 
                ' Projekt aus Showprojekte rausnehmen
                ShowProjekte.Remove(pname)

            End If


            ' die  Variante wird aufgenommen
            ShowProjekte.Add(newProj)

            ' neu zeichnen des Projekts 
            Dim tmpCollection As New Collection
            Call ZeichneProjektinPlanTafel(tmpCollection, newProj.name, tfzeile, phaseList, milestoneList)

            If selectIT Then

                Try
                    CType(appInstance.Worksheets(arrWsNames(3)), Excel.Worksheet).Shapes.Item(newProj.name).Select()
                Catch ex As Exception

                End Try

            End If

        Else
            'Throw New ArgumentException("Projektvariante existiert nicht")
        End If





    End Sub


    ''' <summary>
    ''' Methode trägt alle Projekte aus ImportProjekte in AlleProjekte bzw. Showprojekte ein, sofern die Anzahl mit der myCollection übereinstimmt
    ''' die Projekte werden in der Reihenfolge auf das Board gezeichnet, wie sie in der myCollection aufgeführt sind
    ''' </summary>
    ''' <param name="myCollection"></param>
    ''' <param name="importDate"></param>
    ''' <param name="scenarioName">wenn scenarioName einen wert hat, dann werden für bereits existierende Projekte Varianten mit dem Namen des Szenario-Namens erzeugt </param>
    ''' <remarks></remarks>
    Public Sub importProjekteEintragen(ByVal myCollection As Collection, ByVal importDate As Date, ByVal pStatus As String, _
                                       Optional ByVal scenarioName As String = "")

        Dim hproj As New clsProjekt, cproj As New clsProjekt
        Dim fullName As String, vglName As String
        Dim pname As String


        Dim anzAktualisierungen As Integer, anzNeuProjekte As Integer
        Dim tafelZeile As Integer = 2
        'Dim shpElement As Excel.Shape
        Dim phaseList As New Collection
        Dim milestoneList As New Collection
        Dim wasNotEmpty As Boolean

        If AlleProjekte.Count > 0 Then
            wasNotEmpty = True
        Else
            wasNotEmpty = False
        End If


        Dim differentToPrevious As Boolean = False

        If myCollection.Count <> ImportProjekte.Count Then
            Throw New ArgumentException("keine Übereinstimmung in der Anzahl gültiger/ímportierter Projekte - Abbruch!")
        End If


        anzAktualisierungen = 0
        anzNeuProjekte = 0

        Dim ok As Boolean = True
        ' jetzt werden alle importierten Projekte bearbeitet 
        For Each fullName In myCollection


            ok = True

            Try
                hproj = ImportProjekte.getProject(fullName)
                pname = hproj.name

                ' Änderung tk: ist Filter aktiv ? wenn ja, muss der überprüft werden 
                ' 18.1.15
                If awinSettings.applyFilter Then

                    Dim filter As clsFilter = filterDefinitions.retrieveFilter("Last")
                    If IsNothing(filter) Then
                        ok = True
                    Else
                        ok = filter.doesNotBlock(hproj)
                    End If
                Else
                    ok = True
                End If

            Catch ex As Exception
                Call MsgBox("Projekt " & fullName & " ist kein gültiges Projekt ... es wird ignoriert ...")
                pname = ""
                ok = False
            End Try

            If ok Then

                ' jetzt muss überprüft werden, ob dieses Projekt bereits in AlleProjekte / Showprojekte existiert 
                ' wenn ja, muss es um die entsprechenden Werte dieses Projektes (Status, etc)  ergänzt werden
                ' wenn nein, wird es im Show-Modus ergänzt 

                vglName = calcProjektKey(hproj)
                Try
                    cproj = AlleProjekte.getProject(vglName)
                    anzAktualisierungen = anzAktualisierungen + 1


                    ' hier muss geprüft werden, ob sich die zeitlichen Versionen unterscheiden 
                    ' in einem ersten Schritt werden die Phasen Werte verglichen 
                    ' später soll dann folgen: Ressourcen, Strategischer Fit, Volumen, etc. 


                    ' jetzt wird geprüft , ob die 


                    ' es existiert schon - deshalb müssen alle restlichen Werte aus dem cproj übernommen werden 
                    Try
                        With hproj
                            .farbe = cproj.farbe
                            .Schrift = cproj.Schrift
                            .Schriftfarbe = cproj.Schriftfarbe
                            .earliestStart = cproj.earliestStart
                            .earliestStartDate = cproj.earliestStartDate
                            .Id = vglName & "#" & importDate.ToString
                            .latestStart = cproj.latestStart
                            .latestStartDate = cproj.latestStartDate
                            .leadPerson = cproj.leadPerson
                            .shpUID = cproj.shpUID
                            .StartOffset = 0

                            ' Änderung 28.1.14: bei einem bereits existierenden Projekt muss der Status mitübernommen werden 
                            .Status = cproj.Status ' wird evtl , falls sich Änderungen ergeben haben, noch geändert ...
                            .tfZeile = cproj.tfZeile
                            .timeStamp = importDate
                            .UID = cproj.UID
                            .VorlagenName = cproj.VorlagenName

                            ' im Folgenden werden die Werte dann vom letzten stand übernommen, wenn es keine Werte in 
                            ' der Import datei dafür gab

                            If .StrategicFit = 0 Then
                                .StrategicFit = cproj.StrategicFit
                            End If

                            If .Risiko = 0 Then
                                .Risiko = cproj.Risiko
                            End If

                            If .businessUnit = "" Then
                                .businessUnit = cproj.businessUnit
                            End If

                            If .description = "" Then
                                .description = cproj.description
                            End If

                            If .complexity = 0 Then
                                .complexity = cproj.complexity
                            End If

                            If .volume = 0 Then
                                .volume = cproj.volume
                            End If

                            If cproj.Erloes > 0 Then
                                ' dann soll der alte Wert beibehalten werden 
                                .Erloes = cproj.Erloes
                                If .anzahlRasterElemente = cproj.anzahlRasterElemente And Not IsNothing(cproj.budgetWerte) Then
                                    .budgetWerte = cproj.budgetWerte
                                Else
                                    ' Workaround: 
                                    Dim tmpValue As Integer = hproj.dauerInDays
                                    Call awinCreateBudgetWerte(hproj)
                                End If
                            End If


                            Dim unterschiede As New Collection
                            ' jetzt wird geprüft , ob die beiden Projekte von den Werten her unterschiedlich sind 
                            ' es wird auf absolute Identität geprüft, d.h alleine wenn sich das Startdatum schon verändert gibt es Unterschiede 
                            unterschiede = hproj.listOfDifferences(vglproj:=cproj, absolut:=True, type:=0)
                            If unterschiede.Count > 0 Then
                                ' das heisst, das Projekt hat sich verändert 
                                .diffToPrev = True
                                If .Status = ProjektStatus(1) Then
                                    .Status = ProjektStatus(2)
                                End If
                            Else
                                .diffToPrev = False
                            End If

                        End With
                    Catch ex As Exception
                        ok = False
                        Throw New ArgumentException("Fehler bei Übernahme der Attribute des alten Projektes" & vbLf & ex.Message)

                    End Try


                    Try

                        AlleProjekte.Remove(vglName)
                        If ShowProjekte.contains(hproj.name) Then
                            ShowProjekte.Remove(hproj.name)
                        End If


                    Catch ex1 As Exception
                        Throw New ArgumentException("Fehler beim Update des Projektes " & ex1.Message)
                    End Try


                Catch ex As Exception

                    ' hier existiert das Projekt noch nicht in der AlleProjekte - muss also neu aufgenommen werden 
                    ' es wird auch gleich in showprojekte aufgenommen und in der Plantafel gezeichnet 
                    anzNeuProjekte = anzNeuProjekte + 1


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

                            .Id = vglName & "#" & importDate.ToString
                            ' 5.5.2014 ur: soll nicht wieder auf 0 gesetzt werden, sondern Einstellung beibehalten
                            '.latestStart = 0
                            .latestStartDate = .startDate
                            .leadPerson = " "
                            .shpUID = ""
                            .StartOffset = 0

                            ' ein importiertes Projekt soll normalerweise immer gleich  auf "beauftragt" gesetzt werden; 
                            ' das kann aber jetzt an der aufrufenden Stelle gesetzt werden 
                            ' Inventur: erst mal auf geplant, sonst beauftragt 
                            .Status = pStatus

                            'If DateDiff(DateInterval.Month, .startDate, Date.Now) < -1 Then
                            '    .Status = ProjektStatus(0)
                            'Else
                            '    .Status = ProjektStatus(1)
                            'End If

                            '.tfSpalte = 0
                            .tfZeile = tafelZeile
                            .timeStamp = importDate
                            .UID = cproj.UID

                        End With

                        ' Workaround: 
                        Dim tmpValue As Integer = hproj.dauerInDays
                        Call awinCreateBudgetWerte(hproj)
                        tafelZeile = tafelZeile + 1
                    Catch ex1 As Exception
                        Throw New ArgumentException("Fehler bei Übernahme der Attribute des alten Projektes" & vbLf & ex1.Message)
                    End Try

                End Try

                ' in beiden Fällen - sowohl bei neu wie auch Aktualisierung muss jetzt das Projekt 
                ' sowohl auf der Plantafel eingetragen werden als auch in ShowProjekte und in alleProjekte eingetragen 

                ' bringe das neue Projekt in Showprojekte und in AlleProjekte

                If ok Then

                    Try

                        AlleProjekte.Add(vglName, hproj)
                        ShowProjekte.Add(hproj)

                        ' ggf Bedarfe anzeigen 
                        If roentgenBlick.isOn Then
                            With roentgenBlick
                                Call awinShowNeedsofProject1(mycollection:=.myCollection, type:=.type, projektname:=pname)
                            End With

                        End If

                        ' Änderung tk 18.1.15
                        ' kein Zeichnen - das wird am Schluss komplett gemacht 
                        '' zeichne das neue Shape in der Plan-Tafel 
                        '' wenn bestimmte Projekte beim Suchen nach einem Platz nicht berücksichtigt werden sollen,
                        '' dann müssen sie in einer Collection an ZeichneProjektinPlanTafel übergeben werden 
                        'Dim tmpCollection As New Collection
                        'Call ZeichneProjektinPlanTafel(tmpCollection, pname, hproj.tfZeile, phaseList, milestoneList)

                        '' jetzt müssen die ggf aktuell gezeigten Diagramme neu gezeichnet werden 
                        'Call awinNeuZeichnenDiagramme(2)
                        ' Ende Änderung tk 18.1.15


                    Catch ex As Exception
                        'ur:16.1.2015: Dies ist kein Fehler sondern gewollt: 
                        'Call MsgBox("Fehler bei Eintrag Showprojekte / Import " & hproj.name)
                    End Try

                End If


            End If

        Next

        If ImportProjekte.Count < 1 Then
            Call MsgBox(" es waren keine Projekte zu importieren ...")
        Else
            Dim filterText As String
            If awinSettings.applyFilter Then
                filterText = " (Filter aktiviert)"
            Else
                filterText = " (Filter nicht aktiviert)"
            End If
            Call MsgBox("es wurden " & ImportProjekte.Count & " Projekte bearbeitet!" & filterText & vbLf & vbLf & _
                        anzNeuProjekte.ToString & " neue Projekte" & vbLf & _
                        anzAktualisierungen.ToString & " Projekt-Aktualisierungen")

            ' Änderung tk: jetzt wird das neu gezeichnet 

            If wasNotEmpty Then
                Call awinClearPlanTafel()
            End If

            Call awinZeichnePlanTafel(True)

        End If

        ImportProjekte.Clear()

    End Sub

    ''' <summary>
    ''' überprüft, ob die Eingabe eine Dezimal Zahl > 0 ist 
    ''' </summary>
    ''' <param name="selrange"></param>
    ''' <remarks></remarks>
    Private Sub InputZahlValidationforRange(ByRef selrange As Range)

        With selrange.Validation
            .Delete()
            .Add(Type:=XlDVType.xlValidateDecimal, AlertStyle:=XlDVAlertStyle.xlValidAlertStop, _
                 Operator:=XlFormatConditionOperator.xlGreaterEqual, Formula1:="0")
            .IgnoreBlank = True
            .InCellDropdown = True
            .InputTitle = ""
            .ErrorTitle = ""
            .InputMessage = ""
            .ErrorMessage = "bitte nur Zahlen >= 0 eingeben"
            .ShowInput = True
            .ShowError = True
        End With

    End Sub


    '
    Private Sub InputValidationforRange(ByRef selrange As Range, ByVal stage As Integer, showvalidation As Boolean)

        ' Diese Subroutine erstellt ein Dropdownliste für die Felder von selrange mit den Phasen und Rollen als Auswahl
        ' für stage = 1: Rollen-Kostenarten-Liste

        ' für stage = 2: Phasen-Liste

        Dim inputstr As String = " ", iTitle As String = " ", eMessage As String = " "
        Dim i As Integer
        Dim wasTrue As Boolean

        wasTrue = False


        If appInstance.Application.EnableEvents = True Then
            wasTrue = True
            appInstance.Application.EnableEvents = False
        End If

        If showvalidation = True Then

            If stage = 1 Then
                'Rollen in die Auswahlliste aufnehmen
                inputstr = RoleDefinitions.getRoledef(1).name
                For i = 2 To RoleDefinitions.Count
                    inputstr = inputstr & ";" & RoleDefinitions.getRoledef(i).name
                Next i
                'Kostenarten zur Auswahlliste hinzufügen
                For i = 2 To CostDefinitions.Count - 1
                    inputstr = inputstr & ";" & CostDefinitions.getCostdef(i).name
                Next i
                iTitle = "Rollen/Kostenarten"
                eMessage = "nur Rollen/Kostenarten aus der Liste sind zugelassen ! "
            ElseIf stage = 2 Then
                ' Phasen in die Auswahlliste aufnehmen
                inputstr = PhaseDefinitions.getPhaseDef(1).name
                For i = 2 To PhaseDefinitions.Count
                    inputstr = inputstr & ";" & PhaseDefinitions.getPhaseDef(i).name
                Next
                iTitle = "Phasen"
                eMessage = "nur Phasen aus der Liste sind zugelassen ! "
            End If


            With selrange.Validation
                .Delete()
                .Add(Type:=XlDVType.xlValidateList, AlertStyle:=XlDVAlertStyle.xlValidAlertStop, Operator:= _
                           XlFormatConditionOperator.xlBetween, Formula1:=inputstr)
                .IgnoreBlank = True
                .InCellDropdown = True
                .InputTitle = iTitle
                .ErrorTitle = "Fehler"
                .InputMessage = "bitte auswählen"
                .ErrorMessage = eMessage
                .ShowInput = True
                .ShowError = True
            End With

        Else
            With selrange.Validation
                .Delete()
            End With
            selrange.Locked = True
        End If

        If wasTrue Then
            appInstance.Application.EnableEvents = True
        End If

    End Sub
    ' ''' <summary>
    ' ''' 
    ' ''' </summary>
    ' ''' <param name="constellationName"></param>
    ' ''' <remarks></remarks>
    'Public Sub awinStoreConstellation(ByVal constellationName As String)

    '    '        Dim request As New Request(awinSettings.databaseName)
    '    ' prüfen, ob diese Constellation bereits existiert ..
    '    If projectConstellations.Contains(constellationName) Then

    '        Try
    '            projectConstellations.Remove(constellationName)
    '        Catch ex As Exception

    '        End Try

    '    End If

    '    Dim newC As New clsConstellation
    '    With newC
    '        .constellationName = constellationName
    '    End With

    '    Dim newConstellationItem As clsConstellationItem
    '    For Each kvp As KeyValuePair(Of String, clsProjekt) In ShowProjekte.Liste
    '        newConstellationItem = New clsConstellationItem
    '        With newConstellationItem
    '            .projectName = kvp.Key
    '            .show = True
    '            .Start = kvp.Value.startDate
    '            .variantName = kvp.Value.variantName
    '            .zeile = kvp.Value.tfZeile
    '        End With
    '        newC.Add(newConstellationItem)
    '    Next


    '    Try
    '        projectConstellations.Add(newC)

    '    Catch ex As Exception
    '        Call MsgBox("Fehler bei Add projectConstellations in awinStoreConstellations")
    '    End Try

    '    '' Portfolio in die Datenbank speichern
    '    'If request.pingMongoDb() Then
    '    '    If Not request.storeConstellationToDB(newC) Then
    '    '        Call MsgBox("Fehler beim Speichern der projektConstellation '" & newC.constellationName & "' in die Datenbank")
    '    '    End If
    '    'Else
    '    '    Throw New ArgumentException("Datenbank-Verbindung ist unterbrochen!")
    '    'End If


    'End Sub
    ''' <summary>
    ''' bestimmt die Art des Shapes, ob es ein Projekt, Phasen, Meilenstein, Status oder
    ''' Dependency Shape ist 
    ''' </summary>
    ''' <param name="shape"></param>
    ''' <returns>gibt den Wert gemäß Enumeration PTshty zurück</returns>
    ''' <remarks></remarks>
    Public Function kindOfShape(ByVal shape As Excel.Shape) As Integer

        Dim tmpValue As Integer = -1


        With shape

            If CBool(.HasChart) Then
                tmpValue = -999
            Else

                Try
                    If .AlternativeText.Length > 0 Then
                        tmpValue = CInt(.AlternativeText)
                    End If
                Catch ex As Exception

                End Try

            End If


        End With

        kindOfShape = tmpValue

    End Function

    ''' <summary>
    ''' extrahiert den Projekt-, Phasen- Namen bzw. die Meilenstein Nummer
    ''' je nachdem was als type angegeben wurde 
    ''' </summary>
    ''' <param name="shapeName"></param>
    ''' <param name="type">kann einer der ptshty Enumeration Werte sein</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function extractName(ByVal shapeName As String, ByVal type As Integer) As String

        Dim shpNameParts() As String
        Dim tmpName As String = ""

        shpNameParts = shapeName.Split(New Char() {CChar("#")}, 5)

        If isProjectType(type) Or type = CInt(PTshty.status) Then

            tmpName = shpNameParts(0)

        ElseIf type = PTshty.phaseE Or type = PTshty.phaseN Or type = PTshty.phase1 Then

            tmpName = shpNameParts(1)

        ElseIf type = PTshty.milestoneE Or type = PTshty.milestoneN Then

            ' Änderung tk 17.4. Meilenstein - Name wird genauso extrahiert wie Phasen-Name
            tmpName = shpNameParts(1)

            ' alter Code
            'msNameParts = shpNameParts(2).Split(New Char() {CChar("M")}, 2)
            'tmpName = msNameParts(1)

        End If

        extractName = tmpName

    End Function


    ''' <summary>
    ''' gibt zurück, ob es sich bei dem angegebenen Typ um einen Projekt-Typ handelt
    ''' </summary>
    ''' <param name="type"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function isProjectType(ByVal type As Integer) As Boolean

        If type = PTshty.projektE Or type = PTshty.projektN Or type = PTshty.projektC Or type = PTshty.projektL Then
            isProjectType = True
        Else
            isProjectType = False
        End If

    End Function

    ''' <summary>
    ''' gibt zurück, ob es sich bei dem angegebenen Shape um einen Meilenstein handelt 
    ''' </summary>
    ''' <param name="type"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function isMilestoneType(ByVal type As Integer) As Boolean

        If type = PTshty.milestoneN Or type = PTshty.milestoneE Then
            isMilestoneType = True
        Else
            isMilestoneType = False
        End If

    End Function

    ''' <summary>
    ''' gibt zurück, ob es sich bei dem angegebenen Shape um eine Phase handelt 
    ''' </summary>
    ''' <param name="type"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function isPhaseType(ByVal type As Integer) As Boolean

        If type = PTshty.phase1 Or type = PTshty.phaseE Or type = PTshty.phaseN Then
            isPhaseType = True
        Else
            isPhaseType = False
        End If

    End Function

    ''' <summary>
    ''' gibt true zurück , wenn es sich um ein einzelnes , d.h nicht gruppiertes Projekt-Shape handelt
    ''' false in allen anderen Fällen
    ''' </summary>
    ''' <param name="shpElement"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function isSingleProjectShape(ByVal shpElement As Excel.Shape) As Boolean

        Dim tmpErg As Boolean = False

        With shpElement
            If .AlternativeText = CInt(PTshty.projektL).ToString Then
                tmpErg = True
            ElseIf .AutoShapeType = MsoAutoShapeType.msoShapeRoundedRectangle Then
                If .AlternativeText = CInt(PTshty.projektN).ToString Then
                    tmpErg = True
                End If
            End If
        End With

        isSingleProjectShape = tmpErg

    End Function

    ''' <summary>
    ''' gibt für die Phase zurück, ob Sie in dem angegebenen Zeitraum liegt oder nicht
    ''' wenn kein Timeframe definiert ist, dann wird true zurückgegeben  
    ''' </summary>
    ''' <param name="projektstart">Monat, in dem das Projekt startet</param>
    ''' <param name="relstart">Relativer Monats-Index Phasenstart</param>
    ''' <param name="relEnde">Relative Monats-Index Phasenende</param>
    ''' <param name="von">Monats-Index des linken Randes</param>
    ''' <param name="bis">Monats-Index des rechten Randes</param>
    ''' <returns>true: wenn die Phase diesen Zeitraum berührt
    ''' false: wenn nicht</returns>
    ''' <remarks></remarks>
    Public Function phaseWithinTimeFrame(ByVal projektstart As Integer, ByVal relStart As Integer, ByVal relEnde As Integer, _
                                             ByVal von As Integer, ByVal bis As Integer) As Boolean

        Dim within As Boolean = False

        If von = 0 And bis = 0 Then
            ' wenn kein Zitraum definiert ist, soll true zurückgegeben werden
            within = True
        Else
            If (projektstart + relStart - 1 > bis) Or (projektstart + relEnde - 1 < von) Then
                ' dann liegt die Phase ausserhalb des betrachteten Zeitraums 
                within = False
            Else
                within = True
            End If
        End If


        phaseWithinTimeFrame = within

    End Function

    ''' <summary>
    ''' gibt für das angegebene Datum zurück, ob es in dem angegebenen Zeitraum liegt oder nicht 
    ''' wenn kein Zeitraum definiert ist, wird true zurückgegeben 
    ''' </summary>
    ''' <param name="msDate">Datum</param>
    ''' <param name="von">Monats-Index des linken Randes</param>
    ''' <param name="bis">Monats-Index des rechten Randes</param>
    ''' <returns>true: wenn das Datum innerhalb liegt
    ''' false: sonst</returns>
    ''' <remarks></remarks>
    Public Function milestoneWithinTimeFrame(ByVal msDate As Date, _
                                                 ByVal von As Integer, ByVal bis As Integer) As Boolean

        Dim within As Boolean = False

        If von = 0 And bis = 0 Then
            within = True
        Else
            If DateDiff(DateInterval.Day, StartofCalendar, msDate) >= 0 Then
                If getColumnOfDate(msDate) > bis Or getColumnOfDate(msDate) < von Then
                    within = False
                Else
                    within = True
                End If
            End If
        End If

        milestoneWithinTimeFrame = within

    End Function


    ''' <summary>
    ''' sorgt dafür. daß Projekte immer im gleichen Muster angezeigt werden 
    ''' Erst sortiert nach BU, dann nach ProjektStart-Datum, dann nach Länge  
    ''' </summary>
    ''' <param name="hproj"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function calcKennziffer(ByVal hproj As clsProjekt) As Double

        Dim wertigkeitBU As Integer = 100000
        Dim wertigkeitDate As Double = 100
        Dim wertigkeitLaenge As Double = 0.1
        Dim zwErg As Double = 0.0

        Dim found As Boolean = False
        Dim i As Integer = 1

        While i <= businessUnitDefinitions.Count And Not found

            If businessUnitDefinitions.ElementAt(i - 1).Value.name = hproj.businessUnit Then
                found = True
            Else
                i = i + 1
            End If

        End While

        zwErg = i * wertigkeitBU

        ' Berücksichtigung ProjektstartDatum 
        zwErg = zwErg + DateDiff(DateInterval.Day, StartofCalendar, hproj.startDate) / 30.4 * wertigkeitDate

        ' Berücksichtigung Länge
        zwErg = zwErg + hproj.dauerInDays / 30.4 * wertigkeitLaenge

        calcKennziffer = zwErg

    End Function

    ''' <summary>
    ''' kopiert eine Collection , die Strings enthält 
    ''' </summary>
    ''' <param name="original"></param>
    ''' <param name="kopie"></param>
    ''' <remarks></remarks>
    Public Sub copyCollections(ByVal original As Collection, ByRef kopie As Collection)
        Dim i As Integer
        Dim element As String

        If Not IsNothing(original) Then
            For i = 1 To original.Count
                element = CStr(original.Item(i))
                kopie.Add(element, element)
            Next
        End If

    End Sub

    ''' <summary>
    ''' addiert die Hierarchie hry zur bereits existierenden Super-Hierarchie
    ''' wenn die Super Hierarchie noch leer ist, wird die Rootphase angelegt 
    ''' </summary>
    ''' <param name="superHry"></param>
    ''' <remarks></remarks>
    Public Sub addToSuperHierarchy(ByRef superHry As clsHierarchy, _
                                   ByVal hproj As clsProjekt,
                                   Optional vorlagenIndex As Integer = -1)

        Dim superNode As clsHierarchyNode
        Dim elemID As String
        Dim hry As clsHierarchy = hproj.hierarchy

        ' hier muss später noch ergänzt werden, dass evtl eine Vorlage übergeben werden kann 


        ' ist es eine 
        ' Starten mit dem Rootkey, der ist bei beiden definitiv nur einmal vorhanden und jeweils gleich 

        If superHry.count = 0 Then
            superHry = New clsHierarchy

            superNode = New clsHierarchyNode
            With superNode
                .elemName = elemNameOfElemID(rootPhaseName)
                .indexOfElem = 1 'eigentlich nicht relevant , wird einfach immer auf 1 gesetzt
                .parentNodeKey = ""
                .origName = .elemName
            End With

            superHry.addNode(superNode, rootPhaseName)

        End If

        If IsNothing(hproj) Then
            Exit Sub
        ElseIf IsNothing(hry) Then
            Exit Sub
        ElseIf hry.count <= 1 Then
            Exit Sub
        End If


        ' Schleife über alle Elemente 

        For px As Integer = 1 To hproj.CountPhases

            Dim cphase As clsPhase = hproj.getPhase(px)

            If Not IsNothing(cphase) Then

                elemID = cphase.nameID
                Call addElementToSuperHry(hry, elemID, superHry)

                For mx As Integer = 1 To cphase.countMilestones

                    Dim cMilestone = cphase.getMilestone(mx)
                    If Not IsNothing(cMilestone) Then
                        elemID = cMilestone.nameID
                        Call addElementToSuperHry(hry, elemID, superHry)
                    End If

                Next

            End If



        Next
        'For ix As Integer = 1 To hry.count

        '    hrynode = hry.nodeItem(ix)
        '    curElemID = hry.getIDAtIndex(ix)
        '    curElemName = elemNameOfElemID(curElemID)
        '    isMilestone = elemIDIstMeilenstein(curElemID)
        '    breadcrumb = hry.getBreadCrumb(curElemID)
        '    Dim newBreadcrumb As String = ""

        '    If isMilestone Then
        '        elemIndices = superHry.getMilestoneHryIndices(curElemName, breadcrumb)
        '    Else
        '        elemIndices = superHry.getPhaseHryIndices(curElemName, breadcrumb)
        '    End If



        '    If elemIndices(0) > 0 Then
        '        ' dann wurde ein Element mit dem komplett identischen Breadcrumb gefunden; es ist also gar nichts zu tun 

        '    Else
        '        Dim ptr As Integer = 1
        '        Dim itemExists As Boolean = True
        '        parentNodeID = rootPhaseName
        '        bcItems = breadcrumb.Split((New Char() {CChar("#")}), 30)
        '        Dim anzahlEbenen As Integer = bcItems.Length - 1
        '        Dim lastFoundID As String = rootPhaseName

        '        Do While ptr <= anzahlEbenen And itemExists

        '            tmpElemName = bcItems(ptr)
        '            For i As Integer = 0 To ptr - 1
        '                If i = 0 Then
        '                    newBreadcrumb = "."
        '                Else
        '                    newBreadcrumb = newBreadcrumb & "#" & bcItems(i)
        '                End If

        '            Next

        '            elemIndices = superHry.getPhaseHryIndices(tmpElemName, newBreadcrumb)
        '            If elemIndices(0) > 0 Then
        '                itemExists = True
        '                ptr = ptr + 1
        '                parentNodeID = superHry.nodeItem(elemIndices(0)).parentNodeKey
        '                lastFoundID = superHry.getIDAtIndex(elemIndices(0))
        '            Else
        '                itemExists = False
        '                parentNodeID = lastFoundID
        '            End If

        '        Loop

        '        If ptr > anzahlEbenen Then
        '            parentNodeID = lastFoundID
        '        End If

        '        ' jetzt ist man an der Stelle angelangt, wo die Phase schon nicht mehr existiert 
        '        For i As Integer = ptr To anzahlEbenen
        '            tmpElemName = bcItems(i)
        '            ' lege die Phase an 
        '            superNode = New clsHierarchyNode
        '            With superNode
        '                .elemName = tmpElemName
        '                .parentNodeKey = parentNodeID
        '                .origName = ""
        '                .isMilestone = False ' es handelt sich hier noch um die Hierarchie-Stufen, also Phasen
        '                .indexOfElem = 1 ' eigentlich in diesem Kontext nicht relevant 
        '            End With
        '            elemID = superHry.findUniqueElemKey(tmpElemName, False)
        '            superHry.addNode(superNode, elemID)

        '            ' weiterschalten 
        '            parentNodeID = elemID
        '        Next

        '        ' jetzt muss noch das Element selber angelegt werde 
        '        superNode = New clsHierarchyNode
        '        With superNode
        '            .elemName = curElemName
        '            .parentNodeKey = parentNodeID
        '            .origName = ""
        '            .isMilestone = isMilestone  ' es handelt sich hier noch um die Hierarchie-Stufen, also Phasen
        '            .indexOfElem = 1 ' eigentlich in diesem Kontext nicht relevant 
        '        End With
        '        elemID = superHry.findUniqueElemKey(curElemName, False)
        '        superHry.addNode(superNode, elemID)

        '    End If



        'Next


    End Sub

    ''' <summary>
    ''' trägt das Element mit der übergebenen ElemID in die Super-Hierarchie ein  
    ''' </summary>
    ''' <param name="hry"></param>
    ''' <param name="elemID"></param>
    ''' <param name="superHry"></param>
    ''' <remarks></remarks>
    Private Sub addElementToSuperHry(ByVal hry As clsHierarchy, ByVal elemID As String, ByRef superHry As clsHierarchy)

        Dim hryNode As clsHierarchyNode
        Dim superNode As clsHierarchyNode
        Dim elemIndices() As Integer
        Dim curElemID As String
        Dim curElemName As String
        Dim breadcrumb As String
        Dim isMilestone As Boolean
        Dim parentNodeID As String

        Dim bcItems() As String
        Dim tmpElemName As String



        'hryNode = hry.nodeItem(ix)
        'curElemID = hry.getIDAtIndex(ix)
        hryNode = hry.nodeItem(elemID)
        curElemID = elemID
        curElemName = elemNameOfElemID(curElemID)
        isMilestone = elemIDIstMeilenstein(curElemID)
        breadcrumb = hry.getBreadCrumb(curElemID)

        Dim newBreadcrumb As String = ""

        If isMilestone Then
            elemIndices = superHry.getMilestoneHryIndices(curElemName, breadcrumb)
        Else
            elemIndices = superHry.getPhaseHryIndices(curElemName, breadcrumb)
        End If


        If elemIndices(0) > 0 Then
            ' dann wurde ein Element mit dem komplett identischen Breadcrumb gefunden; es ist also gar nichts zu tun 

        Else
            Dim ptr As Integer = 1
            Dim itemExists As Boolean = True
            parentNodeID = rootPhaseName
            bcItems = breadcrumb.Split((New Char() {CChar("#")}), 30)
            Dim anzahlEbenen As Integer = bcItems.Length - 1
            Dim lastFoundID As String = rootPhaseName

            Do While ptr <= anzahlEbenen And itemExists

                tmpElemName = bcItems(ptr)
                For i As Integer = 0 To ptr - 1
                    If i = 0 Then
                        newBreadcrumb = "."
                    Else
                        newBreadcrumb = newBreadcrumb & "#" & bcItems(i)
                    End If

                Next

                elemIndices = superHry.getPhaseHryIndices(tmpElemName, newBreadcrumb)
                If elemIndices(0) > 0 Then
                    itemExists = True
                    ptr = ptr + 1
                    parentNodeID = superHry.nodeItem(elemIndices(0)).parentNodeKey
                    lastFoundID = superHry.getIDAtIndex(elemIndices(0))
                Else
                    itemExists = False
                    parentNodeID = lastFoundID
                End If

            Loop

            If ptr > anzahlEbenen Then
                parentNodeID = lastFoundID
            End If

            ' jetzt ist man an der Stelle angelangt, wo die Phase schon nicht mehr existiert 
            For i As Integer = ptr To anzahlEbenen
                tmpElemName = bcItems(i)
                ' lege die Phase an 
                superNode = New clsHierarchyNode
                With superNode
                    .elemName = tmpElemName
                    .parentNodeKey = parentNodeID
                    .origName = ""
                    .indexOfElem = 1 ' eigentlich in diesem Kontext nicht relevant 
                End With
                curElemID = superHry.findUniqueElemKey(tmpElemName, False)
                superHry.addNode(superNode, curElemID)

                ' weiterschalten 
                parentNodeID = curElemID
            Next

            ' jetzt muss noch das Element selber angelegt werde 
            superNode = New clsHierarchyNode
            With superNode
                .elemName = curElemName
                .parentNodeKey = parentNodeID
                .origName = ""
                .indexOfElem = 1 ' eigentlich in diesem Kontext nicht relevant 
            End With
            curElemID = superHry.findUniqueElemKey(curElemName, isMilestone)
            superHry.addNode(superNode, curElemID)

        End If

    End Sub

    ''' <summary>
    ''' wird nur zum Aufsetzen von zufälligen Bewertungen in Demo-Szenarien benötigt ... 
    ''' 
    ''' </summary>
    ''' <param name="yellowPercentage">gibt an wieviele Meilensteine gelb bewertet werden sollen</param>
    ''' <param name="redPercentage">gibt an, wieviele Meilensteine rot bewertet werden sollen</param>
    ''' <param name="heute">gibt das Datum an, das als das heutige gelten soll</param> 
    ''' <remarks></remarks>
    Public Sub createInitialRandomBewertungen(ByVal yellowPercentage As Double, ByVal redPercentage As Double, ByVal heute As Date)

        Dim expl As String = "Erläuterung ..."
        Dim redBaseValue As Double = 0.3
        Dim yellowBaseValue As Double = 0.7
        Dim zufall As New Random(10)


        Dim allMilestones As Integer
        Dim redMilestones As Integer
        Dim yellowMilestones As Integer
        Dim greenMilestones As Integer
        Dim firstMS As Integer
        Dim lastMS As Integer
        Dim currentValue As Double
        Dim heuteColumn As Integer = getColumnOfDate(heute)

        For Each kvp As KeyValuePair(Of String, clsProjekt) In ShowProjekte.Liste

            allMilestones = 0
            redMilestones = 0
            yellowMilestones = 0
            greenMilestones = 0

            With kvp.Value
                firstMS = .hierarchy.getIndexOf1stMilestone
                lastMS = .hierarchy.count

                For i As Integer = firstMS To lastMS
                    Dim msID As String = .hierarchy.getIDAtIndex(i)
                    Dim milestone As clsMeilenstein = .getMilestoneByID(msID)
                    Dim msColumn As Integer = getColumnOfDate(milestone.getDate)

                    If msColumn <= heuteColumn + 6 Then

                       

                        currentValue = zufall.NextDouble
                        With milestone
                            If currentValue >= redBaseValue And _
                                currentValue <= redBaseValue + redPercentage Then

                                Dim b As clsBewertung

                                If .bewertungsCount = 0 Then
                                    b = New clsBewertung
                                    b.description = "Erläuterung für die rote Ampel ..."
                                    b.color = awinSettings.AmpelRot
                                    .addBewertung(b)
                                Else
                                    b = .getBewertung(1)
                                    b.description = "Erläuterung für die rote Ampel ..."
                                    b.color = awinSettings.AmpelRot
                                End If
                                
                                If msColumn > heuteColumn Then
                                    redMilestones = redMilestones + 1
                                End If



                            ElseIf currentValue >= yellowBaseValue And _
                                currentValue <= yellowBaseValue + yellowPercentage Then
                                Dim b As clsBewertung

                                If .bewertungsCount = 0 Then
                                    b = New clsBewertung
                                    b.description = "Erläuterung für die gelbe Ampel ..."
                                    b.color = awinSettings.AmpelGelb
                                    .addBewertung(b)
                                Else
                                    b = .getBewertung(1)
                                    b.description = "Erläuterung für die gelbe Ampel ..."
                                    b.color = awinSettings.AmpelGelb
                                End If

                                If msColumn > heuteColumn Then
                                    yellowMilestones = yellowMilestones + 1
                                End If



                            Else
                                Dim b As clsBewertung

                                If .bewertungsCount = 0 Then
                                    b = New clsBewertung
                                    b.description = "aktuell alles i.O.  ..."
                                    b.color = awinSettings.AmpelGruen
                                    .addBewertung(b)
                                Else
                                    b = .getBewertung(1)
                                    b.description = "aktuell alles i.O.  ..."
                                    b.color = awinSettings.AmpelGruen
                                End If

                                If msColumn > heuteColumn Then
                                    greenMilestones = greenMilestones + 1
                                End If

                            End If


                        End With
                    Else
                        ' nichts tun, alles unverändert lassen 
                    End If


                Next

                ' jetzt noch die Ampel-Farbe setzen 
                If redMilestones > 0 Then
                    If redMilestones / greenMilestones > 0.02 Then
                        .ampelErlaeuterung = "Erläuterung des Projektleiters ... "
                        .ampelStatus = 3
                    Else
                        .ampelErlaeuterung = "Erläuterung für gelbe Bewertung (u.a mind. eine rote Ampel) ..."
                        .ampelStatus = 2
                    End If

                ElseIf yellowMilestones > 0 Then
                    If greenMilestones > 0 Then

                        If yellowMilestones / greenMilestones > 0.05 Then
                            .ampelErlaeuterung = "Erläuterung des Projektleiters ... "
                            .ampelStatus = 2
                        Else
                            .ampelErlaeuterung = "aktuell alles i.O ..."
                            .ampelStatus = 1
                        End If
                    Else
                        .ampelErlaeuterung = "Erläuterung des Projektleiters ... "
                        .ampelStatus = 2
                    End If
                Else
                    .ampelErlaeuterung = "aktuell alles i.O ..."
                    .ampelStatus = 1
                End If

            End With

        Next

    End Sub


    ''' <summary>
    ''' es werden zufällig Phasen verschoben, verkürzt bzw. verlängert 
    ''' dadurch werden auch Meilensteine, die in den Phasen sind, vorgezogen oder nach hinten geschoben 
    ''' ausserdem werden dadurch auch Ressourcen durch die proportionale Anpassung weniger / mehr.  
    ''' es werden allerdings nur Phasen verlängert/verkürzt die
    ''' noch nicht beendet sind 
    ''' die bereits begonnen haben bzw. deren Start nicht weiter weg als 2 M ist.  
    ''' </summary>
    ''' <param name="shorterPercentage"></param>
    ''' <param name="longerPercentage"></param>
    ''' <param name="heute"></param>
    ''' <remarks></remarks>
    Public Sub createRandomChanges(ByVal shorterPercentage As Double, ByVal longerPercentage As Double, ByVal heute As Date)

        Dim expl As String = "Erläuterung ..."
        Dim redBaseValue As Double = 0.3
        Dim yellowBaseValue As Double = 0.7
        Dim zufall As New Random(10)

        Dim moveForward As Double = 0.1
        Dim moveBackward As Double = 0.8
        Dim makeItShorter As Double = 0.1
        Dim makeItLonger As Double = 0.8
        Dim currentValue As Double

        Dim cphase As clsPhase
        Dim previousSetting As Boolean = awinSettings.propAnpassRess
        Dim startColumn As Integer, endColumn As Integer
        Dim heuteColumn As Integer = getColumnOfDate(heute)

        ' bisheriges Setting merken 
        awinSettings.propAnpassRess = True

        For Each kvp As KeyValuePair(Of String, clsProjekt) In ShowProjekte.Liste

            With kvp.Value

                For pi As Integer = 1 To .CountPhases
                    cphase = .getPhase(pi)
                    startColumn = getColumnOfDate(cphase.getStartDate)
                    endColumn = getColumnOfDate(cphase.getEndDate)

                    ' wenn die Ausgangsbedingung für nach vorne /hinten verschieben zutrifft: noch nicht gestartet, aber Start weniger als 2 Monate entfernt  
                    If heuteColumn <= startColumn And heuteColumn + 2 >= startColumn Then
                        ' Start nach vorne bzw hinten verschieben
                        currentValue = zufall.NextDouble
                        If currentValue <= moveForward Then
                            Dim anzahlTage As Long = DateDiff(DateInterval.Day, heute, cphase.getStartDate)
                            If anzahlTage > 0 Then
                                Dim newStartOffset As Integer = cphase.startOffsetinDays - CInt(anzahlTage * currentValue)

                            End If


                        ElseIf currentValue >= moveBackward Then

                        End If
                    End If

                    ' wenn die Ausgangsbedingung für verkürzen / verlängern zutrifft: bereits gestartet, aber noch nicht beendet 
                Next

            End With

        Next

        ' altes Setting wiederherstellen 
        awinSettings.propAnpassRess = previousSetting

    End Sub
End Module
