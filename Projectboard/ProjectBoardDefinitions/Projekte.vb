Imports Microsoft.Office.Interop
Imports Microsoft.Office.Interop.Excel
Imports System.Windows.Forms
Imports Microsoft.Office.Core


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
        maxscale = System.Math.Max(10, maxValue * 1.3)

        If maxscale < 100 Then
            maxscale = System.Math.Round(maxscale / 10, MidpointRounding.ToEven) * 10
        Else
            maxscale = System.Math.Round(maxscale / 100, MidpointRounding.ToEven) * 100
        End If

        If maxscale < 10 Then maxscale = 10
        majorUnit = maxscale / 4


        

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

        With appInstance.Worksheets(arrWsNames(3))
            anzDiagrams = .ChartObjects.Count
            '
            ' um welches Diagramm handelt es sich ...
            '
            i = 1
            found = False
            While i <= anzDiagrams And Not found
                Try
                    chtTitle = .ChartObjects(i).Chart.ChartTitle.text
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

                With appInstance.Charts.Add
                    ' remove extra series
                    Do Until .SeriesCollection.Count = 0
                        .SeriesCollection(1).Delete()
                    Loop



                    With .SeriesCollection.NewSeries
                        .name = name1
                        .Interior.color = vergleichsfarbe1
                        .Values = array1
                        .XValues = Xdatenreihe
                        ' Unterschied farblich hervorheben ...
                        For ix = 1 To maxlength
                            If array1(ix - 1) = array2(ix - 1) Then
                                With .Points(ix)
                                    .Interior.color = vergleichsfarbe0
                                End With
                            End If
                        Next
                        .ChartType = Excel.XlChartType.xlColumnClustered
                    End With

                    With .SeriesCollection.NewSeries
                        .name = name2
                        .Interior.color = vergleichsfarbe2
                        .Values = array2
                        .XValues = Xdatenreihe
                        ' Unterschied farblich hervorheben ...
                        For ix = 1 To maxlength
                            If array1(ix - 1) = array2(ix - 1) Then
                                With .Points(ix)
                                    .Interior.color = vergleichsfarbe0
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

                    With .Axes(Excel.XlAxisType.xlCategory)
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

                    With .Axes(Excel.XlAxisType.xlValue)
                        .HasTitle = False
                        .MinimumScale = 0
                        .HasMinorGridlines = False
                        .HasMajorGridlines = True
                        .majorUnit = majorUnit

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
                        .Position = XlConstants.xlTop
                        .Font.Size = 10
                        '.Font.Size = MsoAutoSize.msoAutoSizeTextToFitShape
                    End With

                    .HasTitle = True
                    With .ChartTitle
                        .Text = diagramtitle
                        .font.size = 12
                        '.Font.Size = MsoAutoSize.msoAutoSizeTextToFitShape
                    End With

                    .Location(Where:=XlChartLocation.xlLocationAsObject, Name:=appInstance.Worksheets(arrWsNames(3)).name)
                End With

                ' jetzt kommt die Korrektur der Größe; herausfinden, wieviel Raum die Axis Beschriftung einnimmt ... 
                With .ChartObjects(anzDiagrams + 1)
                    .top = top
                    .height = 2 * height

                    Dim axleft As Double, axwidth As Double
                    If .Chart.HasAxis(Excel.XlAxisType.xlValue) = True Then
                        With .Chart.Axes(Excel.XlAxisType.xlValue)
                            axleft = .left
                            axwidth = .width
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
        maxscale = System.Math.Max(10, maxValue * 1.3)

        If maxscale < 100 Then
            maxscale = System.Math.Round(maxscale / 10, MidpointRounding.ToEven) * 10
        Else
            maxscale = System.Math.Round(maxscale / 100, MidpointRounding.ToEven) * 100
        End If

        If maxscale < 10 Then maxscale = 10
        majorUnit = maxscale / 4




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

        With appInstance.Worksheets(arrWsNames(3))
            anzDiagrams = .ChartObjects.Count
            '
            ' um welches Diagramm handelt es sich ...
            '
            i = 1
            found = False
            While i <= anzDiagrams And Not found
                Try
                    chtTitle = .ChartObjects(i).Chart.ChartTitle.text
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

                With appInstance.Charts.Add
                    ' remove extra series
                    Do Until .SeriesCollection.Count = 0
                        .SeriesCollection(1).Delete()
                    Loop


                    'series
                    With .SeriesCollection.NewSeries
                        .name = "identisch"
                        .Interior.color = vergleichsfarbe0
                        .Values = array0
                        .XValues = Xdatenreihe
                        .ChartType = Excel.XlChartType.xlColumnStacked
                    End With


                    With .SeriesCollection.NewSeries
                        '.name = "mehr"
                        .name = name1
                        .Interior.color = vergleichsfarbe1
                        .Values = array1
                        .XValues = Xdatenreihe
                        .ChartType = Excel.XlChartType.xlColumnStacked
                    End With

                    With .SeriesCollection.NewSeries
                        '.name = "weniger"
                        .name = name2
                        .Interior.color = vergleichsfarbe2
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

                    With .Axes(Excel.XlAxisType.xlCategory)
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

                    With .Axes(Excel.XlAxisType.xlValue)
                        .HasTitle = False
                        .MinimumScale = 0
                        .HasMinorGridlines = False
                        .HasMajorGridlines = True
                        .majorUnit = majorUnit

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
                        .Position = XlConstants.xlTop
                        .Font.Size = 10
                        '.Font.Size = MsoAutoSize.msoAutoSizeTextToFitShape
                    End With

                    .HasTitle = True
                    With .ChartTitle
                        .Text = diagramtitle
                        .font.size = 12
                        '.Font.Size = MsoAutoSize.msoAutoSizeTextToFitShape
                    End With

                    .Location(Where:=XlChartLocation.xlLocationAsObject, Name:=appInstance.Worksheets(arrWsNames(3)).name)
                End With

                ' jetzt kommt die Korrektur der Größe; herausfinden, wieviel Raum die Axis Beschriftung einnimmt ... 
                With .ChartObjects(anzDiagrams + 1)
                    .top = top
                    .height = 2 * height

                    Dim axleft As Double, axwidth As Double
                    If .Chart.HasAxis(Excel.XlAxisType.xlValue) = True Then
                        With .Chart.Axes(Excel.XlAxisType.xlValue)
                            axleft = .left
                            axwidth = .width
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

        With appInstance.Worksheets(arrWsNames(3))
            anzDiagrams = .ChartObjects.Count
            '
            ' um welches Diagramm handelt es sich ...
            '
            i = 1
            found = False
            While i <= anzDiagrams And Not found
                Try
                    chtTitle = .ChartObjects(i).Chart.ChartTitle.text
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

                With appInstance.Charts.Add
                    ' remove extra series
                    Do Until .SeriesCollection.Count = 0
                        .SeriesCollection(1).Delete()
                    Loop


                    'series
                    'With .SeriesCollection.NewSeries
                    '    .name = "identischer Teil"
                    '    .Interior.color = vergleichsfarbe0
                    '    .Values = array0
                    '    .XValues = Xdatenreihe
                    '    .ChartType = Excel.XlChartType.xlColumnStacked
                    'End With

                    With .SeriesCollection.NewSeries
                        .name = name1
                        .Interior.color = vergleichsfarbe1
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

                    With .Axes(Excel.XlAxisType.xlCategory)
                        .HasTitle = True
                        '.MinimumScale = 0
                        With .AxisTitle
                            .Characters.text = "Monate"
                            .Font.Size = 8
                        End With
                    End With

                    With .Axes(Excel.XlAxisType.xlValue)
                        .HasTitle = True
                        '.MinimumScale = 0
                        With .AxisTitle
                            .Characters.text = comparisonItem
                            .Font.Size = 8
                        End With
                    End With

                    .HasLegend = True
                    With .Legend
                        .Position = XlConstants.xlTop
                        .Font.Size = 8
                    End With

                    .HasTitle = True
                    With .ChartTitle
                        .Text = diagramtitle
                        .font.size = 10
                    End With

                    .Location(Where:=XlChartLocation.xlLocationAsObject, Name:=appInstance.Worksheets(arrWsNames(3)).name)
                End With

                With .ChartObjects(anzDiagrams + 1)
                    .top = top
                    .left = left
                    .height = height
                    .width = width
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
    ''' <param name="qualifier"></param>
    ''' <remarks>wenn hier etwas geändert wird, muss auch in updatePhasesBalken geändert werden ... 
    ''' </remarks>
    Public Sub createPhasesBalken(ByVal noColorCollection As Collection, ByVal hproj As clsProjekt, ByRef repObj As Excel.ChartObject, ByVal maxscale As Double, _
                                      ByVal top As Double, ByVal left As Double, ByVal height As Double, ByVal width As Double, ByVal qualifier As String)
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



        Dim formerEE As Boolean = appInstance.EnableEvents
        'Dim formerSU As Boolean = appInstance.ScreenUpdating
        appInstance.EnableEvents = False
        'appInstance.ScreenUpdating = False

        Try
            If qualifier = "Vorlage" Then
                titelTeile(0) = "Vorlage " & hproj.VorlagenName & vbLf
                titelTeile(1) = " "
                kennung = hproj.VorlagenName.Trim & "#Phasen#1"


            ElseIf qualifier = "Beauftragung" Then
                titelTeile(0) = "Beauftragung " & hproj.startDate.ToShortDateString & _
                                     " - " & hproj.startDate.AddDays(hproj.dauerInDays - 1).ToShortDateString & vbLf
                titelTeile(1) = " (" & hproj.timeStamp.ToString & ") "
                kennung = hproj.name.Trim & "Beauftragung" & "#Phasen#1"

            ElseIf qualifier = "letzter Stand" Then
                titelTeile(0) = "letzter Stand " & hproj.startDate.ToShortDateString & _
                                     " - " & hproj.startDate.AddDays(hproj.dauerInDays - 1).ToShortDateString & vbLf
                titelTeile(1) = " (" & hproj.timeStamp.ToString & ") "
                kennung = hproj.name.Trim & "letzter Stand" & "#Phasen#1"

            Else
                titelTeile(0) = hproj.name & " ,  " & hproj.startDate.ToShortDateString & _
                                     " - " & hproj.startDate.AddDays(hproj.dauerInDays - 1).ToShortDateString & vbLf

                titelTeile(1) = " (" & hproj.timeStamp.ToString & ") "
                kennung = hproj.name.Trim & "#Phasen#1"
            End If
        Catch ex As Exception
            titelTeile(0) = hproj.name & vbLf
            titelTeile(1) = " (" & hproj.timeStamp.ToString & ") "
            kennung = hproj.name.Trim & "#Phasen#1"
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
                    If noColorCollection.Contains(.name) Then
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



        With appInstance.Worksheets(arrWsNames(3))
            anzDiagrams = .ChartObjects.Count
            '
            ' um welches Diagramm handelt es sich ...
            '
            i = 1
            found = False
            While i <= anzDiagrams And Not found

                Try
                    If .ChartObjects(i).Name = kennung Then
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
                repObj = .chartobjects(i)
            Else

                With appInstance.Charts.Add
                    ' remove extra series
                    Do Until .SeriesCollection.Count = 0
                        .SeriesCollection(1).Delete()
                    Loop

                    'Aufbau der Series 

                    With .SeriesCollection.NewSeries

                        For i = 0 To anzPhasen - 1
                            mdatenreihe(i) = tdatenreihe1(i) / 365 * 12
                        Next
                        .name = "null1"
                        .Interior.colorindex = -4142
                        .Values = mdatenreihe
                        .XValues = Xdatenreihe
                        .HasDataLabels = False

                        For px = 1 To anzPhasen

                            With .Points(px)
                                If tdatenreihe1(px - 1) < 90 Then
                                    .HasDataLabel = False
                                Else
                                    .HasDataLabel = True
                                    .Datalabel.Text = hproj.startDate.AddDays(tdatenreihe1(px - 1)).ToShortDateString
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

                    With .SeriesCollection.NewSeries

                        For i = 0 To anzPhasen - 1
                            mdatenreihe(i) = tdatenreihe2(i) / 365 * 12
                        Next
                        .name = "Phasen Zeitraum"
                        .Values = mdatenreihe
                        .XValues = Xdatenreihe

                        .HasDataLabels = True
                        .DataLabels.Font.Size = awinSettings.fontsizeItems
                        .DataLabels.Position = Excel.XlDataLabelPosition.xlLabelPositionCenter

                        For i = 1 To anzPhasen
                            With .Points(i)
                                .Interior.Color = valueColor(i - 1)

                                If mdatenreihe(i - 1) <= 3 Then
                                    .Datalabel.Text = tdatenreihe2(i - 1).ToString
                                Else
                                    .Datalabel.Text = tdatenreihe2(i - 1).ToString & " Tage"
                                End If

                            End With
                        Next


                        .ChartType = Excel.XlChartType.xlBarStacked
                    End With

                    With .SeriesCollection.NewSeries

                        .name = "null2"
                        .Interior.colorindex = -4142
                        .Values = tdatenreihe3
                        .XValues = Xdatenreihe

                        .HasDataLabels = True
                        .DataLabels.Font.Size = awinSettings.fontsizeItems + 2
                        .DataLabels.Position = Excel.XlDataLabelPosition.xlLabelPositionInsideBase

                        Dim bis As Integer
                        For px = 1 To anzPhasen

                            With .Points(px)

                                bis = tdatenreihe1(px - 1) + tdatenreihe2(px - 1)
                                .Datalabel.Text = hproj.startDate.AddDays(bis - 1).ToShortDateString

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
                    .ChartTitle.font.size = awinSettings.fontsizeTitle
                    .ChartTitle.Format.TextFrame2.TextRange.Characters(titelTeilLaengen(0) + 1, _
                                                                        titelTeilLaengen(1)).Font.Size = awinSettings.fontsizeLegend
                    .Location(Where:=XlChartLocation.xlLocationAsObject, Name:=appInstance.Worksheets(arrWsNames(3)).name)
                End With

                ' jetzt kommt die Korrektur der Größe; herausfinden, wieviel Raum die Axis Beschriftung einnimmt ... 
                With .ChartObjects(anzDiagrams + 1)

                    .chart.ChartTitle.Format.TextFrame2.TextRange.Characters(titelTeilLaengen(0) + 1, _
                                                                   titelTeilLaengen(1)).Font.Size = awinSettings.fontsizeLegend
                    .top = top
                    .height = (anzPhasen - 1) * 20 + 90

                    Dim axCleft As Double, axCwidth As Double
                    If .Chart.HasAxis(Excel.XlAxisType.xlCategory) = True Then
                        With .Chart.Axes(Excel.XlAxisType.xlCategory)
                            axCleft = .left
                            axCwidth = .width
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

                    .Name = kennung


                End With

                repObj = .ChartObjects(anzDiagrams + 1)

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
    ''' <param name="minscale"></param>
    ''' <param name="maxscale"></param>
    ''' <remarks></remarks>
    Public Sub updatePhasesBalken(ByVal hproj As clsProjekt, ByRef chtobj As Excel.ChartObject, ByVal minscale As Double, maxscale As Double)
        Dim diagramTitle As String

        Dim anzPhasen As Integer

        Dim plenInDays As Integer
        Dim i As Integer
        Dim Xdatenreihe() As String
        Dim valueColor() As Object
        Dim tdatenreihe1() As Double, mdatenreihe() As Double, tdatenreihe2() As Double, tdatenreihe3() As Double
        Dim titelTeile(1) As String
        Dim titelTeilLaengen(1) As Integer




        Dim pname As String = hproj.name



        Dim formerEE As Boolean = appInstance.EnableEvents
        'Dim formerSU As Boolean = appInstance.ScreenUpdating
        appInstance.EnableEvents = False
        'appInstance.ScreenUpdating = False



        titelTeile(0) = hproj.name & " ,  " & hproj.startDate.ToShortDateString & _
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

        If anzPhasen < 1 Then
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



        With chtobj.Chart
            ' remove extra series
            Do Until .SeriesCollection.Count = 0
                .SeriesCollection(1).Delete()
            Loop

            'Aufbau der Series 

            With .SeriesCollection.NewSeries

                For i = 0 To anzPhasen - 1
                    mdatenreihe(i) = tdatenreihe1(i) / 365 * 12
                Next
                .name = "null1"
                .Interior.colorindex = -4142
                .Values = mdatenreihe
                .XValues = Xdatenreihe
                .HasDataLabels = False

                For px = 1 To anzPhasen

                    With .Points(px)
                        If tdatenreihe1(px - 1) = 0 Then
                            .HasDataLabel = False
                        Else
                            .HasDataLabel = True
                            .Datalabel.Text = hproj.startDate.AddDays(tdatenreihe1(px - 1)).ToShortDateString
                            .DataLabel.Font.Size = awinSettings.fontsizeItems + 2
                            .DataLabel.Position = Excel.XlDataLabelPosition.xlLabelPositionInsideEnd
                        End If

                    End With

                Next

                .ChartType = Excel.XlChartType.xlBarStacked
            End With

            With .SeriesCollection.NewSeries

                For i = 0 To anzPhasen - 1
                    mdatenreihe(i) = tdatenreihe2(i) / 365 * 12
                Next
                .name = "Phasen Zeitraum"
                .Values = mdatenreihe
                .XValues = Xdatenreihe

                .HasDataLabels = True
                .DataLabels.Font.Size = awinSettings.fontsizeItems
                .DataLabels.Position = Excel.XlDataLabelPosition.xlLabelPositionCenter

                For i = 1 To anzPhasen
                    With .Points(i)
                        .Interior.Color = valueColor(i - 1)
                        If mdatenreihe(i - 1) <= 3 Then
                            .Datalabel.Text = tdatenreihe2(i - 1).ToString
                        Else
                            .Datalabel.Text = tdatenreihe2(i - 1).ToString & " Tage"
                        End If
                    End With
                Next

                
                .ChartType = Excel.XlChartType.xlBarStacked
            End With

            With .SeriesCollection.NewSeries

                .name = "null2"
                .Interior.colorindex = -4142
                .Values = tdatenreihe3
                .XValues = Xdatenreihe

                .HasDataLabels = True
                .DataLabels.Font.Size = awinSettings.fontsizeItems + 2
                .DataLabels.Position = Excel.XlDataLabelPosition.xlLabelPositionInsideBase

                Dim bis As Integer
                For px = 1 To anzPhasen

                    With .Points(px)

                        bis = tdatenreihe1(px - 1) + tdatenreihe2(px - 1)
                        .Datalabel.Text = hproj.startDate.AddDays(bis - 1).ToShortDateString

                    End With

                Next

                .ChartType = Excel.XlChartType.xlBarStacked

            End With


            With .Axes(Excel.XlAxisType.xlValue)

                .MinimumScale = minscale / 365 * 12
                .MaximumScale = CInt(maxscale / 365 * 12) + 3

            End With



            .HasTitle = True
            .ChartTitle.Text = diagramTitle
            .ChartTitle.Font.Size = awinSettings.fontsizeTitle
            .ChartTitle.Format.TextFrame2.TextRange.Characters(titelTeilLaengen(0) + 1, _
                    titelTeilLaengen(1)).Font.Size = awinSettings.fontsizeLegend
            '.ChartTitle.Font.Size = awinSettings.fontsizeTitle
        End With


        'Call awinScrollintoView()
        appInstance.EnableEvents = formerEE
        'appInstance.ScreenUpdating = formerSU


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

        Dim werteB(beauftragung.Dauer - 1) As Double
        Dim werteL(lastPlan.Dauer - 1) As Double
        Dim werteC(hproj.Dauer - 1) As Double

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
        titelTeile(1) = pname & vbLf
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
            maxColumn = .Start + .Dauer - 1
        End With

        With beauftragung
            If maxColumn < .Start + .Dauer - 1 Then
                maxColumn = .Start + .Dauer - 1
            End If
        End With

        With lastPlan
            If maxColumn < .Start + .Dauer - 1 Then
                maxColumn = .Start + .Dauer - 1
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

            endeIX = System.Math.Min(heuteColumn - 1, .Start + .Dauer - 1)
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
            endeIX = System.Math.Min(heuteColumn - 1, .Start + .Dauer - 1)
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
            endeIX = System.Math.Min(heuteColumn - 1, .Start + .Dauer - 1)
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



        With appInstance.Worksheets(arrWsNames(3))
            anzDiagrams = .ChartObjects.Count
            '
            ' um welches Diagramm handelt es sich ...
            '
            i = 1
            found = False
            While i <= anzDiagrams And Not found
                Dim chtTitle As String
                Try
                    chtTitle = .ChartObjects(i).Chart.ChartTitle.text
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
                reportObj = .ChartObjects(i)
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

                chtobj = .Chartobjects(anzDiagrams + 1)
                chtobj.Name = pname & "#" & kennung & "#" & "1"


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
                            .name = "Baseline"
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
                    .name = "Current"
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
            Try
                lastPlan = projekthistorie.ElementAtorBefore(vgl)
            Catch ex As Exception
                Throw New ArgumentException("es gibt keinen Stand vorher")
            End Try


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

        Dim werteB(beauftragung.Dauer - 1) As Double
        Dim werteL(lastPlan.Dauer - 1) As Double
        Dim werteC(hproj.Dauer - 1) As Double

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
        titelTeile(1) = pname & vbLf
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
            maxColumn = .Start + .Dauer - 1
        End With

        With beauftragung
            If maxColumn < .Start + .Dauer - 1 Then
                maxColumn = .Start + .Dauer - 1
            End If
        End With

        With lastPlan
            If maxColumn < .Start + .Dauer - 1 Then
                maxColumn = .Start + .Dauer - 1
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
                If i >= .Start And i <= .Start + .Dauer - 1 Then
                    tdatenreiheB(i - minColumn) = sumB + werteB(i - .Start)
                    sumB = tdatenreiheB(i - minColumn)
                Else
                    tdatenreiheB(i - minColumn) = sumB
                End If
            End With

            With lastPlan
                If i >= .Start And i <= .Start + .Dauer - 1 Then
                    tdatenreiheL(i - minColumn) = sumL + werteL(i - .Start)
                    sumL = tdatenreiheL(i - minColumn)
                Else
                    tdatenreiheL(i - minColumn) = sumL
                End If
            End With

            With hproj
                If i >= .Start And i <= .Start + .Dauer - 1 Then
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

        With appInstance.Worksheets(arrWsNames(3))
            anzDiagrams = .ChartObjects.Count
            '
            ' um welches Diagramm handelt es sich ...
            '
            i = 1
            found = False
            While i <= anzDiagrams And Not found
                Dim chtTitle As String
                Try
                    chtTitle = .ChartObjects(i).Chart.ChartTitle.text
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
                reportObj = .ChartObjects(i)
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

                chtobj = .Chartobjects(anzDiagrams + 1)
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
                            .name = "Baseline (" & beauftragung.timeStamp.ToString("d") & ")"
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
                    .name = "Current (" & hproj.timeStamp.ToString("d") & ")"
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
    Public Sub createMsTrendAnalysisOfProject(ByRef hproj As clsProjekt, ByRef repObj As Object, ByRef myCollection As Collection, _
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

        titelTeile(0) = "Meilenstein Trend-Analyse " & pname & vbLf
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
            Throw New Exception("es gibt noch keinen Trend")
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


        With appInstance.Worksheets(arrWsNames(3))
            anzDiagrams = .ChartObjects.Count
            '
            ' um welches Diagramm handelt es sich ...
            '
            i = 1
            found = False
            While i <= anzDiagrams And Not found
                Try
                    chtTitle = .ChartObjects(i).Chart.ChartTitle.text
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

                repObj = .ChartObjects(i)
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

                chtobj = .Chartobjects(anzDiagrams + 1)
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
                    msName = myCollection.Item(ms)

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

                            colorIndex = DateDiff(DateInterval.Second, tmpdatenreihe(qx).Date, tmpdatenreihe(qx))
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
                            .Name = drawnMilestones.ToString & " - " & msName
                            .ChartType = Excel.XlChartType.xlLineMarkers
                            .Interior.Color = awinSettings.AmpelNichtBewertet
                            .Values = tdatenreihe
                            .XValues = Xdatenreihe
                            .HasDataLabels = False
                            .MarkerStyle = Excel.XlMarkerStyle.xlMarkerStyleCircle
                            .MarkerForegroundColor = awinSettings.AmpelNichtBewertet
                            .MarkerBackgroundColor = awinSettings.AmpelNichtBewertet

                            With .Format.Line
                                .Visible = MsoTriState.msoTrue
                                .ForeColor.RGB = awinSettings.AmpelNichtBewertet
                                .DashStyle = MsoLineDashStyle.msoLineDashDot
                            End With
                        End With


                        For px = 1 To tdatenreihe.Length
                            With CType(.SeriesCollection(drawnMilestones).Points(px), Point)
                                .Interior.Color = ampelfarben(px - 1)
                                .MarkerStyle = Excel.XlMarkerStyle.xlMarkerStyleCircle
                                .MarkerForegroundColor = ampelfarben(px - 1)
                                .MarkerBackgroundColor = ampelfarben(px - 1)
                                .MarkerSize = 10

                                ' Schreiben des ersten Planungs-Standes
                                If px = 1 Then

                                    ' wenn es der Wert aus dem Vormonat ist: einen kleineren Marker zeichnen 
                                    If prevValueTaken(px - 1) Then
                                        .MarkerSize = 5
                                    End If

                                    ' wenn der Meilenstein zum zeitpunkt des Planungs-Standes bereits in der Vergangenheit lag, wird er auch so markiert
                                    If milestoneReached(px - 1) Then
                                        .MarkerStyle = Excel.XlMarkerStyle.xlMarkerStyleSquare
                                    End If

                                    .HasDataLabel = True
                                    If anzMilestones > 1 Then
                                        .DataLabel.Text = drawnMilestones.ToString & " - " & tmpdatenreihe(px - 1).ToShortDateString
                                    Else
                                        .DataLabel.Text = msName & vbLf & tmpdatenreihe(px - 1).ToShortDateString
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
                                        .MarkerSize = 5
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

                                        End Try

                                        Try
                                            drawnDates.Add(tmpdatenreihe(px - 1).Date, tmpdatenreihe(px - 1))
                                        Catch ex As Exception

                                        End Try


                                    End If
                                End If

                                ' Schreiben des letzten Planungs-Standes
                                If px > 1 And px = tdatenreihe.Length Then

                                    ' wenn es der Wert aus dem Vormonat ist: einen kleineren Marker zeichnen 
                                    If prevValueTaken(px - 1) Then
                                        .MarkerSize = 5
                                    End If

                                    ' wenn der Meilenstein zum zeitpunkt des Planungs-Standes bereits in der Vergangenheit lag, wird er auch so markiert
                                    If milestoneReached(px - 1) Then
                                        .MarkerStyle = Excel.XlMarkerStyle.xlMarkerStyleSquare
                                    End If

                                    .HasDataLabel = True
                                    If anzMilestones > 1 Then
                                        .DataLabel.Text = drawnMilestones.ToString & " - " & tmpdatenreihe(px - 1).ToShortDateString
                                    Else
                                        .DataLabel.Text = msName & vbLf & tmpdatenreihe(px - 1).ToShortDateString
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
                spread = DateDiff(DateInterval.Day, tmpMinScale, tmpMaxScale) / 10
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
                        .Position = Excel.Constants.xlTop
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
    Public Sub createRessBalkenOfProject(ByRef hproj As clsProjekt, ByRef repObj As Object, ByVal auswahl As Integer, _
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

        Dim formerEE As Boolean = appInstance.EnableEvents
        'Dim formerSU As Boolean = appInstance.ScreenUpdating
        appInstance.EnableEvents = False
        'appInstance.ScreenUpdating = False


        Dim pname As String = hproj.name

        If auswahl = 1 Then
            titelTeile(0) = "Ressourcen-Bedarf " & zE & vbLf & pname & vbLf
            titelTeilLaengen(0) = titelTeile(0).Length
            titelTeile(1) = " (" & hproj.timeStamp.ToString & ") "
            titelTeilLaengen(1) = titelTeile(1).Length
            diagramTitle = titelTeile(0) & titelTeile(1)
            kennung = "Personalbedarf"
        ElseIf auswahl = 2 Then
            titelTeile(0) = "Personalkosten (T€)" & vbLf & pname & vbLf
            titelTeilLaengen(0) = titelTeile(0).Length
            titelTeile(1) = " (" & hproj.timeStamp.ToString & ") "
            titelTeilLaengen(1) = titelTeile(1).Length
            diagramTitle = titelTeile(0) & titelTeile(1)
            diagramTitle = titelTeile(0) & titelTeile(1)
            kennung = "Personalkosten"
        Else
            diagramTitle = "--- (T€)" & vbLf & pname
            'kennung = "Gesamtkosten"
        End If



        '
        ' hole die Projektdauer
        '
        With hproj
            plen = .Dauer
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
        With appInstance.Worksheets(arrWsNames(3))
            anzDiagrams = .ChartObjects.Count
            '
            ' um welches Diagramm handelt es sich ...
            '
            i = 1
            found = False
            While i <= anzDiagrams And Not found
                Try
                    chtTitle = .ChartObjects(i).Chart.ChartTitle.text
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
                repObj = .ChartObjects(i)
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

                chtobj = .Chartobjects(anzDiagrams + 1)
                chtobj.Name = pname & "#" & kennung & "#" & "1"


            End If

            With chtobj.Chart

                .ChartTitle.Format.TextFrame2.TextRange.Characters(titelTeilLaengen(0) + 1, _
                    titelTeilLaengen(1)).Font.Size = awinSettings.fontsizeLegend

                For r = 1 To anzRollen
                    roleName = ErgebnisListeR.Item(r)
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
                    With .Chart.Axes(Excel.XlAxisType.xlValue)
                        axleft = .left
                        axwidth = .width
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

    '
    ' Prozedur updated die Ressourcen Struktur des Projektes im Chart chtobj 
    '
    ' Auswahl = 1 : Diagramm zeigt Mann-Monate 
    ' Auswahl = 2 : Diagramm zeigt Personal-Kosten  
    ' Kennung Phasen, Personalbedarf, Personalkosten, Sonstige Kosten, Gesamtkosten, Strategie, Ergebnis

    Public Sub updateRessBalkenOfProject(ByRef hproj As clsProjekt, ByRef chtobj As Excel.ChartObject, ByVal auswahl As Integer, _
                                         ByVal minscale As Double, ByVal maxscale As Double)

        Dim kennung As String = " "
        Dim diagramTitle As String = " "

        Dim plen As Integer
        Dim i As Integer
        Dim Xdatenreihe() As String
        Dim tdatenreihe() As Double
        Dim hsum() As Double, gesamt_summe As Double
        Dim anzRollen As Integer
        Dim pkIndex As Integer = CostDefinitions.Count
        Dim pstart As Integer
        Dim ErgebnisListeR As New Collection
        Dim roleName As String
        Dim zE As String = "(" & awinSettings.kapaEinheit & ")"
        Dim titelTeile(1) As String
        Dim titelTeilLaengen(1) As Integer

        Dim formerEE As Boolean = appInstance.EnableEvents
        appInstance.EnableEvents = False
        'Dim formerSU As Boolean = appInstance.ScreenUpdating
        'appInstance.ScreenUpdating = False


        Dim pname As String = hproj.name

        If auswahl = 1 Then
            titelTeile(0) = "Ressourcen-Bedarf " & zE & vbLf & pname & vbLf
            titelTeilLaengen(0) = titelTeile(0).Length
            titelTeile(1) = " (" & hproj.timeStamp.ToString & ") "
            titelTeilLaengen(1) = titelTeile(1).Length
            diagramTitle = titelTeile(0) & titelTeile(1)
            kennung = "Personalbedarf"
        ElseIf auswahl = 2 Then
            titelTeile(0) = "Personalkosten (T€)" & vbLf & pname & vbLf
            titelTeilLaengen(0) = titelTeile(0).Length
            titelTeile(1) = " (" & hproj.timeStamp.ToString & ") "
            titelTeilLaengen(1) = titelTeile(1).Length
            diagramTitle = titelTeile(0) & titelTeile(1)
            kennung = "Personalkosten"
        Else
            diagramTitle = "--- (T€)" & vbLf & pname
            'kennung = "Gesamtkosten"
        End If



        '
        ' hole die Projektdauer
        '
        With hproj
            plen = .Dauer
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
            MsgBox("keine Ressourcen-Bedarfe definiert")
            Exit Sub
        End If


        ReDim Xdatenreihe(plen - 1)
        ReDim tdatenreihe(plen - 1)


        ReDim hsum(anzRollen - 1)

        For i = 1 To plen
            Xdatenreihe(i - 1) = StartofCalendar.AddMonths(pstart + i - 2).ToString("MMM yy")
        Next i

        gesamt_summe = 0


        With chtobj.Chart

            ' remove extra series
            Do Until .SeriesCollection.Count = 0
                .SeriesCollection(1).Delete()
            Loop

            For r = 1 To anzRollen
                roleName = ErgebnisListeR.Item(r)
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

            .ChartTitle.Text = diagramTitle
            .ChartTitle.Font.Size = awinSettings.fontsizeTitle
            .ChartTitle.Format.TextFrame2.TextRange.Characters(titelTeilLaengen(0) + 1, _
                    titelTeilLaengen(1)).Font.Size = awinSettings.fontsizeLegend

        End With





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

    Public Sub createCostBalkenOfProject(ByRef hproj As clsProjekt, ByRef repObj As Object, ByVal auswahl As Integer, _
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


        Dim ErgebnisListeK As Collection

        Dim formerEE As Boolean = appInstance.EnableEvents
        'Dim formerSU As Boolean = appInstance.ScreenUpdating
        appInstance.EnableEvents = False
        'appInstance.ScreenUpdating = False


        Dim pname As String = hproj.name
        If auswahl = 1 Then

            titelTeile(0) = "Sonstige Kosten T€" & vbLf & pname & vbLf
            titelTeilLaengen(0) = titelTeile(0).Length
            titelTeile(1) = " (" & hproj.timeStamp.ToString & ") "
            titelTeilLaengen(1) = titelTeile(1).Length
            diagramTitle = titelTeile(0) & titelTeile(1)
            kennung = "Sonstige Kosten"
        Else
            titelTeile(0) = "Gesamtkosten T€" & vbLf & pname & vbLf
            titelTeilLaengen(0) = titelTeile(0).Length
            titelTeile(1) = " (" & hproj.timeStamp.ToString & ") "
            titelTeilLaengen(1) = titelTeile(1).Length
            diagramTitle = titelTeile(0) & titelTeile(1)
            kennung = "Gesamtkosten"
        End If


        '
        ' hole die Projektdauer
        '
        With hproj
            plen = .Dauer
            pstart = .Start
        End With
        '
        ' hole die Anzahl Kostenarten, die in diesem Projekt vorkommen
        '
        ErgebnisListeK = hproj.getUsedKosten
        anzKostenarten = ErgebnisListeK.Count

        If anzKostenarten = 0 Then
            MsgBox("keine Kosten-Bedarfe definiert")
            appInstance.EnableEvents = formerEE
            'appInstance.ScreenUpdating = formerSU
            Exit Sub
        End If


        ReDim Xdatenreihe(plen - 1)
        ReDim tdatenreihe(plen - 1)


        If auswahl = 1 Then
            ReDim hsum(anzKostenarten - 1)
        Else
            ReDim hsum(anzKostenarten) ' weil jetzt die berechneten Personalkosten dazu kommen
        End If


        For i = 1 To plen
            Xdatenreihe(i - 1) = StartofCalendar.AddMonths(pstart + i - 2).ToString("MMM yy")
        Next i


        Dim ik As Integer = 1 ' wird für die Unterscheidung benötigt, ob mit Personal-Kosten oder ohne 
        gesamt_summe = 0
        With appInstance.Worksheets(arrWsNames(3))
            anzDiagrams = .ChartObjects.Count
            '
            ' um welches Diagramm handelt es sich ...
            '
            i = 1
            found = False
            While i <= anzDiagrams And Not found
                Try
                    chtTitle = .ChartObjects(i).Chart.ChartTitle.text
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
                repObj = .ChartObjects(i)
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

                chtobj = .Chartobjects(anzDiagrams + 1)
                chtobj.Name = pname & "#" & kennung & "#" & "1"


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
                    costname = ErgebnisListeK.Item(k)
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
                    With .Chart.Axes(Excel.XlAxisType.xlValue)
                        axleft = .left
                        axwidth = .width
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
    ''' aktualisiert das Auslastungs Chart 
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


        Dim kennung As String
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
            titelTeile(0) = summentitel10
        Else
            titelTeile(0) = summentitel11
        End If

        titelTeilLaengen(0) = titelTeile(0).Length + 1
        titelTeile(1) = textZeitraum(showRangeLeft, showRangeRight)
        titelTeilLaengen(1) = titelTeile(1).Length
        diagramTitle = titelTeile(0) & vbLf & titelTeile(1)
        kennung = titelTeile(0)



        With appInstance.Worksheets(arrWsNames(3))


            Dim tmpValues(1) As Double



            With chtobj.Chart
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

                .HasTitle = True
                .ChartTitle.Text = diagramTitle
                .ChartTitle.Font.Size = awinSettings.fontsizeTitle
                .ChartTitle.Format.TextFrame2.TextRange.Characters(titelTeilLaengen(0) + 1, _
                    titelTeilLaengen(1)).Font.Size = awinSettings.fontsizeLegend

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
    ''' <param name="minscale">Angabe für Minimum-Scale</param>
    ''' <param name="maxscale">Angabe für Maximum-Scale</param>
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
    Public Sub updateCostBalkenOfProject(ByRef hproj As clsProjekt, ByRef chtobj As Excel.ChartObject, ByVal auswahl As Integer, _
                                             ByVal minscale As Double, ByVal maxscale As Double)

        Dim kennziffer As Integer = 3
        Dim plen As Integer
        Dim i As Integer
        Dim Xdatenreihe() As String
        Dim tdatenreihe() As Double
        Dim hsum() As Double, gesamt_summe As Double
        Dim anzKostenarten As Integer
        Dim costname As String
        Dim pkIndex As Integer = CostDefinitions.Count
        Dim pstart As Integer
        Dim diagramTitle As String
        Dim titelTeile(1) As String
        Dim titelTeilLaengen(1) As Integer


        Dim pname As String = hproj.name
        If auswahl = 1 Then
            titelTeile(0) = "Sonstige Kosten T€" & vbLf & pname & vbLf
            titelTeilLaengen(0) = titelTeile(0).Length
            titelTeile(1) = " (" & hproj.timeStamp.ToString & ") "
            titelTeilLaengen(1) = titelTeile(1).Length
            diagramTitle = titelTeile(0) & titelTeile(1)

        Else
            titelTeile(0) = "Gesamtkosten T€" & vbLf & pname & vbLf
            titelTeilLaengen(0) = titelTeile(0).Length
            titelTeile(1) = " (" & hproj.timeStamp.ToString & ") "
            titelTeilLaengen(1) = titelTeile(1).Length
            diagramTitle = titelTeile(0) & titelTeile(1)

        End If


        Dim ErgebnisListeK As Collection

        Dim formerEE As Boolean = appInstance.EnableEvents
        appInstance.EnableEvents = False
        'Dim formerSU As Boolean = appInstance.ScreenUpdating
        'appInstance.ScreenUpdating = False

        '
        ' hole die Projektdauer
        '
        With hproj
            plen = .Dauer
            pstart = .Start
        End With
        '
        ' hole die Anzahl Kostenarten, die in diesem Projekt vorkommen
        '
        ErgebnisListeK = hproj.getUsedKosten
        anzKostenarten = ErgebnisListeK.Count

        If anzKostenarten = 0 Then
            MsgBox("keine Kosten-Bedarfe definiert")
            Exit Sub
        End If


        ReDim Xdatenreihe(plen - 1)
        ReDim tdatenreihe(plen - 1)


        If auswahl = 1 Then
            ReDim hsum(anzKostenarten - 1)
        Else
            ReDim hsum(anzKostenarten) ' weil jetzt die berechneten Personalkosten dazu kommen
        End If


        For i = 1 To plen
            Xdatenreihe(i - 1) = StartofCalendar.AddMonths(pstart + i - 2).ToString("MMM yy")
        Next i


        Dim ik As Integer = 1 ' wird für die Unterscheidung benötigt, ob mit Personal-Kosten oder ohne 
        gesamt_summe = 0


        With chtobj.Chart

            ' remove extra series
            Do Until .SeriesCollection.Count = 0
                .SeriesCollection(1).Delete()
            Loop

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
                    '.ChartType = Excel.XlChartType.xlColumnStacked
                End With
            End If

            For k = 1 To anzKostenarten
                costname = ErgebnisListeK.Item(k)
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
                    '.ChartType = Excel.XlChartType.xlColumnStacked
                End With

            Next k

            'With .Axes(Excel.XlAxisType.xlValue)
            '    .MaximumScale = maxscale
            '    .MinimumScale = minscale

            '    'With .AxisTitle
            '    '    .Characters.text = "Kosten"
            '    '    .Font.Size = 8
            '    'End With
            'End With
            .ChartTitle.Text = diagramTitle
            .ChartTitle.Font.Size = awinSettings.fontsizeTitle
            .ChartTitle.Format.TextFrame2.TextRange.Characters(titelTeilLaengen(0) + 1, _
                    titelTeilLaengen(1)).Font.Size = awinSettings.fontsizeLegend
        End With



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
    Public Sub createAuslastungsDetailPie(ByRef repObj As Object, ByVal auswahl As Integer, _
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
            chtobjname = getKennung("pf", PTpfdk.UeberAuslastung, myCollection)
        Else
            chtobjname = getKennung("pf", PTpfdk.Unterauslastung, myCollection)
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
            titelTeile(0) = summentitel10 & " (" & awinSettings.kapaEinheit & ")"
        Else
            titelTeile(0) = summentitel11 & " (" & awinSettings.kapaEinheit & ")"
        End If

        titelTeilLaengen(0) = titelTeile(0).Length + 1
        titelTeile(1) = textZeitraum(showRangeLeft, showRangeRight)
        titelTeilLaengen(1) = titelTeile(1).Length
        diagramTitle = titelTeile(0) & vbLf & titelTeile(1)
        kennung = titelTeile(0)



        With appInstance.Worksheets(arrWsNames(3))
            anzDiagrams = .ChartObjects.Count
            '
            ' um welches Diagramm handelt es sich ...
            '
            Dim i As Integer = 1
            Dim found As Boolean = False
            Dim chtTitle As String
            While i <= anzDiagrams And Not found
                Try
                    'chtTitle = .ChartObjects(i).Chart.ChartTitle.text
                    chtTitle = .ChartObjects(i).name
                Catch ex As Exception
                    chtTitle = " "
                End Try

                If chtTitle = chtobjname Then
                    found = True

                Else
                    i = i + 1
                End If

            End While

            If found Then
                'Call MsgBox("Chart wird bereits angezeigt ...")
                appInstance.EnableEvents = formerEE
                repObj = .ChartObjects(i)
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


            repObj = .ChartObjects(anzDiagrams + 1)

            ' jetzt muss die letzte Position des Diagramms gespeichert werden , wenn es nicht aus der Reporting Engine 
            ' aufgerufen wurde
            If Not calledfromReporting Then

                Dim prcDiagram As New clsDiagramm

                ' Anfang Event Handling für Chart 
                Dim prcChart As New clsEventsPrcCharts
                prcChart.PrcChartEvents = .ChartObjects(anzDiagrams + 1).Chart
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

    Public Sub createRessPieOfProject(ByRef hproj As clsProjekt, ByRef repObj As Object, ByVal auswahl As Integer, _
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


        Dim ErgebnisListeR As Collection
        Dim formerEE As Boolean = appInstance.EnableEvents
        appInstance.EnableEvents = False


        '
        ' hole die Projektdauer
        '
        With hproj
            plen = .Dauer
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
            roleName = ErgebnisListeR.Item(r + 1)
            Xdatenreihe(r) = roleName
            If auswahl = 1 Then
                tdatenreihe(r) = Math.Round(hproj.getRessourcenBedarf(roleName).Sum)
            Else
                tdatenreihe(r) = Math.Round(hproj.getPersonalKosten(roleName).Sum)
            End If

        Next r

        If auswahl = 1 Then
            titelTeile(0) = "Personalbedarf (" & tdatenreihe.Sum.ToString("#####.") & zE & ")" & vbLf & pname & vbLf
            titelTeilLaengen(0) = titelTeile(0).Length
            titelTeile(1) = "(" & hproj.timeStamp.ToString & ") "
            titelTeilLaengen(1) = titelTeile(1).Length
            diagramTitle = titelTeile(0) & titelTeile(1)
            kennung = "Personalbedarf"
        Else
            titelTeile(0) = "Personalkosten (" & tdatenreihe.Sum.ToString("#####.") & " T€)" & vbLf & pname & vbLf
            titelTeilLaengen(0) = titelTeile(0).Length
            titelTeile(1) = "(" & hproj.timeStamp.ToString & ") "
            titelTeilLaengen(1) = titelTeile(1).Length
            diagramTitle = titelTeile(0) & titelTeile(1)

            kennung = "Personalkosten"
        End If


        With appInstance.Worksheets(arrWsNames(3))
            anzDiagrams = .ChartObjects.Count
            '
            ' um welches Diagramm handelt es sich ...
            '
            Dim i As Integer = 1
            Dim found As Boolean = False
            Dim chtTitle As String
            While i <= anzDiagrams And Not found
                Try
                    chtTitle = .ChartObjects(i).Chart.ChartTitle.text
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
                repObj = .ChartObjects(i)
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
                        roleName = ErgebnisListeR.Item(r)
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
                    .Name = pname & "#" & kennung & "#" & "2"
                    .top = top
                    .left = left
                    .height = height
                    .width = width
                End With
            End If


            repObj = .ChartObjects(anzDiagrams + 1)



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

        Dim kennung As String
        Dim zE As String = awinSettings.kapaEinheit & " "
        Dim titelTeile(1) As String
        Dim titelTeilLaengen(1) As Integer


        Dim ErgebnisListeR As Collection


        Dim formerEE As Boolean = appInstance.EnableEvents
        appInstance.EnableEvents = False


        '
        ' hole die Projektdauer
        '
        With hproj
            plen = .Dauer
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
            roleName = ErgebnisListeR.Item(r + 1)
            Xdatenreihe(r) = roleName

            If auswahl = 1 Then
                tdatenreihe(r) = Math.Round(hproj.getRessourcenBedarf(roleName).Sum)
            Else
                tdatenreihe(r) = Math.Round(hproj.getPersonalKosten(roleName).Sum / 10) * 10
            End If

        Next r


        If auswahl = 1 Then
            titelTeile(0) = "Personalbedarf (" & tdatenreihe.Sum.ToString("####.#") & zE & ")" & vbLf & pname & vbLf
            titelTeilLaengen(0) = titelTeile(0).Length
            titelTeile(1) = "(" & hproj.timeStamp.ToString & ") "
            titelTeilLaengen(1) = titelTeile(1).Length
            diagramTitle = titelTeile(0) & titelTeile(1)

            kennung = "Personalbedarf"
        Else
            titelTeile(0) = "Personalkosten (" & tdatenreihe.Sum.ToString("####.#") & " T€)" & vbLf & pname & vbLf
            titelTeilLaengen(0) = titelTeile(0).Length
            titelTeile(1) = "(" & hproj.timeStamp.ToString & ") "
            titelTeilLaengen(1) = titelTeile(1).Length
            diagramTitle = titelTeile(0) & titelTeile(1)

            kennung = "Personalkosten"
        End If



        With chtobj.Chart
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
                roleName = ErgebnisListeR.Item(r)
                With .SeriesCollection(1).Points(r)
                    .Interior.color = RoleDefinitions.getRoledef(roleName).farbe
                    .DataLabel.Font.Size = awinSettings.fontsizeItems
                End With
            Next r

            '.HasLegend = True
            'With .Legend
            '    .Position = Excel.Constants.xlTop
            '    .Font.Size = awinSettings.fontsizeItems
            'End With
            '.HasTitle = True
            .ChartTitle.Text = diagramTitle
            .ChartTitle.Font.Size = awinSettings.fontsizeTitle
            .ChartTitle.Format.TextFrame2.TextRange.Characters(titelTeilLaengen(0) + 1, _
                    titelTeilLaengen(1)).Font.Size = awinSettings.fontsizeLegend
            '.ChartTitle.Font.Size = awinSettings.fontsizeTitle
            '.Location(Where:=XlChartLocation.xlLocationAsObject, Name:=appInstance.Worksheets(arrWsNames(3)).name)
        End With





        'Call awinScrollintoView()
        appInstance.EnableEvents = formerEE
        'appInstance.ScreenUpdating = True


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

    Public Sub createCostPieOfProject(ByRef hproj As clsProjekt, ByRef repObj As Object, ByVal auswahl As Integer, _
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


        Dim ErgebnisListeK As Collection

        Dim formerEE As Boolean = appInstance.EnableEvents
        appInstance.EnableEvents = False
        'appInstance.ScreenUpdating = False


        '
        ' hole die Projektdauer
        '
        With hproj
            plen = .Dauer
            pstart = .Start
        End With
        '
        ' hole die Anzahl Kostenarten, die in diesem Projekt vorkommen
        '
        ErgebnisListeK = hproj.getUsedKosten
        anzKostenarten = ErgebnisListeK.Count

        If anzKostenarten = 0 Then
            appInstance.EnableEvents = formerEE
            Throw New Exception("keine Kosten-Bedarfe definiert")
        End If

        If auswahl = 1 Then
            ' Alle Sonstigen Kostenarten 
            ReDim tdatenreihe(anzKostenarten - 1)
            ReDim Xdatenreihe(anzKostenarten - 1)
        Else
            ' alle Kostenarten - inkl Personalkosten 
            ReDim tdatenreihe(anzKostenarten)
            ReDim Xdatenreihe(anzKostenarten)
            'Xdatenreihe(0) = "Personal-Kosten"
            'tdatenreihe(0) = hsum(0)
        End If


        For k = 0 To anzKostenarten - 1
            costname = ErgebnisListeK.Item(k + 1)
            Xdatenreihe(k) = costname
            tdatenreihe(k) = Math.Round(hproj.getKostenBedarf(costname).Sum)
        Next k

        If auswahl = 2 Then
            Xdatenreihe(anzKostenarten) = "Personal-Kosten"
            tdatenreihe(anzKostenarten) = Math.Round(hproj.getAllPersonalKosten.Sum)
        End If

        If auswahl = 1 Then
            titelTeile(0) = "Sonstige Kosten (" & tdatenreihe.Sum.ToString("#####.") & " T€)" & vbLf & pname & vbLf
            titelTeilLaengen(0) = titelTeile(0).Length
            titelTeile(1) = " (" & hproj.timeStamp.ToString & ") "
            titelTeilLaengen(1) = titelTeile(1).Length
            diagramTitle = titelTeile(0) & titelTeile(1)

            kennung = "Sonstige Kosten"
        Else
            titelTeile(0) = "Gesamtkosten (" & tdatenreihe.Sum.ToString("#####.") & " T€)" & vbLf & pname & vbLf
            titelTeilLaengen(0) = titelTeile(0).Length
            titelTeile(1) = " (" & hproj.timeStamp.ToString & ") "
            titelTeilLaengen(1) = titelTeile(1).Length
            diagramTitle = titelTeile(0) & titelTeile(1)

            kennung = "Gesamtkosten"
        End If

        If tdatenreihe.Sum = 0.0 Then
            appInstance.EnableEvents = formerEE
            Throw New Exception("Summe sonstige Kosten ist Null")
        Else
            With appInstance.Worksheets(arrWsNames(3))
                anzDiagrams = .ChartObjects.Count

                '
                ' um welches Diagramm handelt es sich ...
                '
                Dim i As Integer = 1
                Dim found As Boolean = False
                Dim chtTitle As String
                While i <= anzDiagrams And Not found
                    Try
                        chtTitle = .ChartObjects(i).Chart.ChartTitle.text
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
                    repObj = .ChartObjects(i)
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
                                costname = ErgebnisListeK.Item(k + 1)
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
                        .Name = pname & "#" & kennung & "#" & "2"
                        .top = top
                        .left = left
                        .height = height
                        .width = width
                    End With

                    repObj = .ChartObjects(anzDiagrams + 1)
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

        Dim kennung As String
        Dim titelTeile(1) As String
        Dim titelTeilLaengen(1) As Integer



        Dim ErgebnisListeK As Collection


        Dim formerEE As Boolean = appInstance.EnableEvents
        appInstance.EnableEvents = False
        'appInstance.ScreenUpdating = False




        '
        ' hole die Projektdauer
        '
        With hproj
            plen = .Dauer
            pstart = .Start
        End With
        '
        ' hole die Anzahl Kostenarten, die in diesem Projekt vorkommen
        '
        ErgebnisListeK = hproj.getUsedKosten
        anzKostenarten = ErgebnisListeK.Count

        If anzKostenarten = 0 Then
            MsgBox("keine Kosten-Bedarfe definiert")
            Exit Sub
        End If

        If auswahl = 1 Then
            ' Alle Sonstigen Kostenarten 
            ReDim tdatenreihe(anzKostenarten - 1)
            ReDim Xdatenreihe(anzKostenarten - 1)
        Else
            ' alle Kostenarten - inkl Personalkosten 
            ReDim tdatenreihe(anzKostenarten)
            ReDim Xdatenreihe(anzKostenarten)
            'Xdatenreihe(0) = "Personal-Kosten"
            'tdatenreihe(0) = hsum(0)
        End If


        For k = 0 To anzKostenarten - 1
            costname = ErgebnisListeK.Item(k + 1)
            Xdatenreihe(k) = costname
            tdatenreihe(k) = hproj.getKostenBedarf(costname).Sum
        Next k

        If auswahl = 2 Then
            Xdatenreihe(anzKostenarten) = "Personal-Kosten"
            tdatenreihe(anzKostenarten) = hproj.getAllPersonalKosten.Sum
        End If

        If auswahl = 1 Then
            titelTeile(0) = "Sonstige Kosten (" & tdatenreihe.Sum.ToString("####.#") & " T€)" & vbLf & pname & vbLf
            titelTeilLaengen(0) = titelTeile(0).Length
            titelTeile(1) = " (" & hproj.timeStamp.ToString & ") "
            titelTeilLaengen(1) = titelTeile(1).Length
            diagramTitle = titelTeile(0) & titelTeile(1)
            kennung = "Sonstige Kosten"
        Else
            titelTeile(0) = "Gesamtkosten (" & tdatenreihe.Sum.ToString("####.#") & " T€)" & vbLf & pname & vbLf
            titelTeilLaengen(0) = titelTeile(0).Length
            titelTeile(1) = " (" & hproj.timeStamp.ToString & ") "
            titelTeilLaengen(1) = titelTeile(1).Length
            diagramTitle = titelTeile(0) & titelTeile(1)

            kennung = "Gesamtkosten"
        End If


        With chtobj.Chart
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
                    costname = ErgebnisListeK.Item(k + 1)
                    With .SeriesCollection(1).Points(k + 1)
                        .Interior.color = CostDefinitions.getCostdef(costname).farbe
                        .DataLabel.Font.Size = 10

                    End With
                End If

            Next k

            '.HasLegend = True
            'With .Legend
            '    .Position = Excel.Constants.xlTop
            '    .Font.Size = awinSettings.fontsizeItems
            'End With
            '.HasTitle = True
            .ChartTitle.Text = diagramTitle
            .ChartTitle.Font.Size = awinSettings.fontsizeTitle
            .ChartTitle.Format.TextFrame2.TextRange.Characters(titelTeilLaengen(0) + 1, _
                    titelTeilLaengen(1)).Font.Size = awinSettings.fontsizeLegend
            '.ChartTitle.Font.Size = awinSettings.fontsizeTitle
            '.Location(Where:=XlChartLocation.xlLocationAsObject, Name:=appInstance.Worksheets(arrWsNames(3)).name)
        End With

        'Call awinScrollintoView()
        appInstance.EnableEvents = formerEE
        'appInstance.ScreenUpdating = True


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
    Public Sub createTrendKPI(ByRef repObj As Object, ByVal top As Double, left As Double, height As Double, width As Double)

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

        diagramTitle = "Planungs-Historie Kennzahlen " & vbLf & pname


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


        With appInstance.Worksheets(arrWsNames(3))
            anzDiagrams = .ChartObjects.Count
            '
            ' um welches Diagramm handelt es sich ...
            '
            i = 1
            found = False
            While i <= anzDiagrams And Not found
                Try
                    chtTitle = .ChartObjects(i).Chart.ChartTitle.text
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
                repObj = .ChartObjects(i)
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

                chtobj = .Chartobjects(anzDiagrams + 1)

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
    Public Sub createTrendSfit(ByRef repObj As Object, ByVal top As Double, left As Double, height As Double, width As Double)

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

        diagramTitle = "Planungs-Historie strategischer Fit & Risiko: " & vbLf & pname


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


        With appInstance.Worksheets(arrWsNames(3))
            anzDiagrams = .ChartObjects.Count
            '
            ' um welches Diagramm handelt es sich ...
            '
            i = 1
            found = False
            While i <= anzDiagrams And Not found
                Try
                    chtTitle = .ChartObjects(i).Chart.ChartTitle.text
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
                repObj = .ChartObjects(i)
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

                chtobj = .Chartobjects(anzDiagrams + 1)

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
        plen = hproj.Dauer
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

        With appInstance.Worksheets(arrWsNames(3))
            anzDiagrams = .ChartObjects.Count
            '
            ' um welches Diagramm handelt es sich ...
            '
            i = 1
            found = False
            While i <= anzDiagrams And Not found
                Try
                    chtTitle = .ChartObjects(i).Chart.ChartTitle.text
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
                        With .Chart.Axes(Excel.XlAxisType.xlValue)
                            axleft = .left
                            axwidth = .width
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
    ''' </summary>
    ''' <param name="hproj">das Projekt</param>
    ''' <param name="reportObj">
    ''' nimmt den Verweis auf das generierte Chart auf; 
    ''' wird für das Reporting benötigt 
    ''' </param>
    ''' <remarks></remarks>
    Public Sub createProjektErgebnisCharakteristik2(ByRef hproj As clsProjekt, ByRef reportObj As Object)

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



        '
        ' hole die Projektdauer
        '
        With hproj
            plen = .Dauer
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
        kennung = pname & "#Ergebnis#1"


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






        With appInstance.Worksheets(arrWsNames(3))
            anzDiagrams = .ChartObjects.Count

            '
            ' um welches Diagramm handelt es sich ...
            '
            i = 1
            found = False
            While i <= anzDiagrams And Not found

                'Try
                '    chtTitle = .ChartObjects(i).Chart.ChartTitle.text
                'Catch ex As Exception
                '    chtTitle = " "
                'End Try


                If kennung = .ChartObjects(i).name Then
                    found = True
                Else
                    i = i + 1
                End If

            End While


            Dim currentWert As Double
            If found Then
                reportObj = .ChartObjects(i)
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

                reportObj = .ChartObjects(anzDiagrams + 1)


            End If


        End With





    End Sub

    Public Sub updateProjektErgebnisCharakteristik2(ByRef hproj As clsProjekt, ByRef chtobj As Excel.ChartObject)


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

        Dim formerEE As Boolean = appInstance.EnableEvents
        appInstance.EnableEvents = False



        '
        ' hole die Projektdauer
        '
        With hproj
            plen = .Dauer
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
        kennung = pname & "#Ergebnis#1"

        If projektErgebnis < 0 Then
            minscale = System.Math.Round(projektErgebnis, mode:=MidpointRounding.ToEven)
        Else
            minscale = 0
        End If



        Dim currentWert As Double



        Dim valueCrossesNull As Boolean = False

        With chtobj.Chart
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
                    If minscale < 0 Then
                        .MinimumScale = System.Math.Round((minscale - 1), mode:=MidpointRounding.ToEven)
                    Else
                        .MinimumScale = 0
                    End If
                End With
            Catch ex As Exception

            End Try

            .HasTitle = True
            .ChartTitle.Text = diagramTitle
            .ChartTitle.Font.Size = awinSettings.fontsizeTitle
            .ChartTitle.Format.TextFrame2.TextRange.Characters(titelTeilLaengen(0) + 1, _
                titelTeilLaengen(1)).Font.Size = awinSettings.fontsizeLegend

        End With

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
    Public Sub TrageivProjektein(ByVal pname As String, ByVal vorlagenName As String, ByVal startdate As Date, ByVal erloes As Double, _
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
        Dim heute As Date = Now
        Dim key As String = pname & "#"

        newprojekt = True

        '
        ' ein neues Projekt wird als Objekt angelegt ....
        '

        hproj = New clsProjekt

        Try
            Projektvorlagen.getProject(vorlagenName).CopyTo(hproj)
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
                plen = .Dauer
                pcolor = .farbe
            End With
        Catch ex As Exception
            Call MsgBox(ex.Message)
            Exit Sub
        End Try

        ' Anpassen der Daten für die Termine 
        ' wenn Samstag oder Sonntag, dann auf den Freitag davor legen   


        Dim cphase As clsPhase

        Dim resultDate As Date

        For p = 1 To hproj.CountPhases
            cphase = hproj.getPhase(p)
            For r = 1 To cphase.CountResults

                With cphase.getResult(r)
                    resultDate = .getDate
                    If resultDate.DayOfWeek = DayOfWeek.Saturday Then
                        .offset = .offset - 1
                    ElseIf resultDate.DayOfWeek = DayOfWeek.Sunday Then
                        .offset = .offset - 2
                    End If
                End With

            Next
        Next



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


        Call ZeichneProjektinPlanTafel(pname:=pname, tryzeile:=0, showresults:=False)


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
    Public Sub erstelleÍnventurProjekt(ByRef hproj As clsProjekt, ByVal pname As String, ByVal vorlagenName As String, ByVal startdate As Date, ByVal erloes As Double, _
                          ByVal tafelZeile As Integer, ByVal sfit As Double, ByVal risk As Double, _
                          ByVal volume As Double, ByVal complexity As Double, ByVal businessUnit As String, ByVal description As String)
        Dim newprojekt As Boolean
        Dim pStatus As String = ProjektStatus(0)
        Dim zeile As Integer = tafelZeile
        Dim spalte As Integer = getColumnOfDate(startdate)
        'Dim plen As Integer
        'Dim pcolor As Object
        Dim heute As Date = Now
        Dim key As String = pname & "#"

        newprojekt = True

        '
        ' ein neues Projekt wird als Objekt angelegt ....
        '


        Try
            Projektvorlagen.getProject(vorlagenName).CopyTo(hproj)
        Catch ex As Exception
            Call MsgBox("es gibt keine entsprechende Vorlage ..")
            Exit Sub
        End Try


        Try
            With hproj
                .name = pname
                .getPhase(1).name = pname
                .VorlagenName = vorlagenName
                .startDate = startdate
                .earliestStartDate = .startDate.AddMonths(.earliestStart)
                .latestStartDate = .startDate.AddMonths(.latestStart)
                .Status = ProjektStatus(0)
                .StrategicFit = sfit
                .Risiko = risk
                .volume = volume
                .complexity = complexity
                .businessUnit = businessUnit
                .description = description
                .tfZeile = tafelZeile
                .Erloes = erloes
                '.tfSpalte = start
                'plen = .Dauer
                'pcolor = .farbe
                'If erloes <= 0 Then
                '    .Erloes = System.Math.Round(.getGesamtKostenBedarf.Sum * (1 + .risikoKostenfaktor))
                'Else
                '    .Erloes = erloes
                'End If

            End With
        Catch ex As Exception
            Throw New Exception("in erstelle InventurProjekte: " & ex.Message)
        End Try



        '
        ' Ende Objekt Anlage
        '


    End Sub

    '
    '
    '
    ''' <summary>
    ''' löscht das angegebene Projekt mit Name vprojektName
    ''' </summary>
    ''' <param name="vprojektname"></param>
    ''' <param name="firstCall">
    ''' gibt an , ob es der erste Aufruf war
    ''' wenn ja, kommt erst der Bestätigungs-Dialog 
    ''' wenn nein, wird ohne Aufforderung zur Bestätigung gelöscht 
    ''' </param>
    ''' <remarks></remarks>
    Public Sub awinDeleteChartorProject(ByVal vprojektname As String, ByVal firstCall As Boolean)

        Dim abstand As Integer
        Dim returnValue As DialogResult
        Dim bestaetigeLoeschen As New frmconfirmDeletePrj

        enableOnUpdate = False

        If firstCall Then
            returnValue = bestaetigeLoeschen.ShowDialog

            If returnValue = DialogResult.OK Then
                Call awinLoescheProjekt(vprojektname)
                Call awinClkReset(abstand)

                ' ein Projekt wurde gelöscht  - typus = 3
                Call awinNeuZeichnenDiagramme("3")
            Else
                Throw New ArgumentException("Abbruch")
            End If

        Else
            Call awinLoescheProjekt(vprojektname)
            Call awinClkReset(abstand)

            ' ein Projekt wurde gelöscht  - typus = 3
            Call awinNeuZeichnenDiagramme("3")

        End If



        enableOnUpdate = True

    End Sub

    Public Sub awinDeleteChart(ByRef chtobj As ChartObject)
        Dim chtTitle As String
        Dim found As Boolean


        found = False

        Try
            chtTitle = chtobj.Chart.ChartTitle.Text
        Catch ex As Exception
            chtTitle = " "
        End Try


        If istCockpitDiagramm(chtobj) And TypOfCockpitChart(chtobj) >= 0 Then
            Call awinLoescheCockpitCharts(TypOfCockpitChart(chtobj))

        Else
            chtobj.Delete()

            ' jetzt in DiagrammList suchen ...
            ' Änderung 18.10.13 nicht mehr löschen, weil in der Diagrammliste jetzt die letzte Position des Diagrammes gespeichert wird 
            'i = 1
            'While i <= DiagramList.Count And Not found
            '    If (chtTitle Like (DiagramList.getDiagramm(i).DiagrammTitel & "*")) And _
            '                      (DiagramList.getDiagramm(i).isCockpitChart = False) Then
            '        DiagramList.Remove(i)
            '        found = True
            '    Else
            '        i = i + 1
            '    End If
            'End While

        End If

    End Sub

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
        If ShowProjekte.Liste.ContainsKey(pname) Then


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

                Call ZeichneProjektinPlanTafel(pname, zeile, False)

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
        If ShowProjekte.Liste.ContainsKey(pname) Then


            Try
                hproj = ShowProjekte.getProject(pname)
                With hproj
                    zeile = .tfZeile
                    .Status = ProjektStatus(0)
                    .timeStamp = Date.Now
                End With

                Call ZeichneProjektinPlanTafel(pname, zeile, False)

            Catch ex As Exception
                Call MsgBox(" Fehler in Zurücknahme Beauftragung " & pname & " , Modul: awinCancelBeauftragung")
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
        If ShowProjekte.Liste.ContainsKey(pname) Then

            ' Shape wird gelöscht - ausserdem wird der Verweis in hproj auf das Shape gelöscht 
            Call clearProjektinPlantafel(pname)


            Try
                hproj = ShowProjekte.getProject(pname)
                'NoShowProjekte.Add(hproj)
            Catch ex As Exception
                Call MsgBox(" Fehler in NoShow " & pname & " , Modul: NoShowProject")
                Exit Sub
            End Try


            ShowProjekte.Remove(pname)

            Dim abstand As Integer ' eigentlich nur Dummy Variable, wird aber in Tabelle2 benötigt ...
            Call awinClkReset(abstand)

            ' ein Projekt wurde gelöscht bzw aus Showprojekte entfernt  - typus = 3
            Call awinNeuZeichnenDiagramme("3")



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
                If Not listeRollen.Contains(listeTemp.Item(i)) Then
                    listeRollen.Add(listeTemp.Item(i))
                End If
            Catch ex As Exception

            End Try

        Next

        listeKosten = hproj.getUsedKosten
        listeTemp = cproj.getUsedKosten

        For i = 1 To listeTemp.Count
            Try
                If Not listeKosten.Contains(listeTemp.Item(i)) Then
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
        width = hproj.Dauer * boxWidth + 10


        Dim hname As String, mEinheit As String = awinSettings.kapaEinheit


        ' jetzt werden für alle  Rollenbedarfe, sofern unterschiedlich die Diagramme gezeichnet ... 
        Try
            For i = 1 To listeRollen.Count
                hname = listeRollen.Item(i)
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
                hname = listeKosten.Item(i)
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

        Dim vname As String, phaseName As String
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
                    hvproj.CopyTo(cproj)
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
        ' und nicht gleichzeitig Namen der Projektvorlagen sind 

        For Each cphase In hproj.Liste

            Try
                If Not Projektvorlagen.Liste.ContainsKey(cphase.name) Then
                    Try
                        ergListe.Add(cphase.name, cphase.name)
                    Catch ex As Exception

                    End Try
                End If
            Catch ex1 As Exception

            End Try

        Next

        ' in cproj könnten ja Phasen auftauchen, die in hproj nicht drin sind ...
        For Each cphase In cproj.Liste

            Try
                If Not Projektvorlagen.Liste.ContainsKey(cphase.name) Then
                    Try
                        ergListe.Add(cphase.name, cphase.name)
                    Catch ex As Exception

                    End Try
                End If
            Catch ex1 As Exception

            End Try

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
            phaseName = ergListe.Item(i)
            Xdatenreihe(i - 1) = phaseName
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
            phaseName = ergListe.Item(i)

            Try
                With hproj.getPhase(phaseName)
                    p1(0) = .startOffsetinDays
                    p1(1) = .startOffsetinDays + .dauerInDays - 1
                End With

                Try
                    With cproj.getPhase(phaseName)
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
                    With cproj.getPhase(phaseName)
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




        With appInstance.Worksheets(arrWsNames(3))
            anzDiagrams = .ChartObjects.Count
            '
            ' um welches Diagramm handelt es sich ...
            '
            i = 1
            found = False
            While i <= anzDiagrams And Not found
                Try
                    chtTitle = .ChartObjects(i).Chart.ChartTitle.text
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
                        With .Chart.Axes(Excel.XlAxisType.xlCategory)
                            axCleft = .left
                            axCwidth = .width
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
                                              ByRef liste2 As SortedList(Of Date, String))
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
            plen = .Dauer
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

        With appInstance.Worksheets(arrWsNames(3))
            anzDiagrams = .ChartObjects.Count
            '
            ' um welches Diagramm handelt es sich ...
            '
            i = 1
            found = False
            While i <= anzDiagrams And Not found
                Try
                    chtTitle = .ChartObjects(i).Chart.ChartTitle.text
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
                        roleName = ErgebnisListeR.Item(r)
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
                        With .Chart.Axes(Excel.XlAxisType.xlValue)
                            axleft = .left
                            axwidth = .width
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
                        roleName = ErgebnisListeR.Item(r)
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





    '
    ' zeigt einen Callout an mit den Infos zum Projekt: Name, Anfang, Ende , Dauer
    '
    Public Sub awinShowCallout(ByVal farbindex As Object)

        'Dim top As Double, left As Double, height As Double, width As Double
        Dim startdate As Date


        'Dim textzeile As String

        Call DeleteStartMarkers()

        startdate = StartofCalendar


        'diff = ClickPosition(1, 2) - 1
        'vondate = startdate.AddMonths(diff)
        'von = vondate.ToString("MMM yy")

        'diff = ClickPosition(1, 2) + projektLaenge - 2
        'bisdate = startdate.AddMonths(diff)
        'bis = bisdate.ToString("MMM yy")

        '
        ' die Position des Callouts wird ausgerechnet ...
        '

        Call MsgBox("muss noch implementiert werden - die alte Vorgehensweise siehe auskommentiert im Folgenden ")
        'top = (ClickPosition(1, 1) - 1) * 15 - 31
        'If top < 0 Then
        '    top = 2
        'End If

        'left = (ClickPosition(1, 2) + 2) * boxWidth + 16


        'height = 50
        'width = 150

        'textzeile = selectedProjects(1) & Chr(10) & von & " - " & bis


        'Call awinDrawCallout(textzeile, farbindex, top, left, width, height)


    End Sub

    Public Sub DeleteStartMarkers()
        Dim shp As Excel.Shape

        For Each shp In appInstance.ActiveSheet.Shapes
            With shp
                'If .AutoShapeType = MsoAutoShapeType.msoShapeLineCallout3 Or .AutoShapeType = MsoAutoShapeType.msoShapeIsoscelesTriangle Then
                If .AutoShapeType = MsoAutoShapeType.msoShapeIsoscelesTriangle Then
                    .Delete()
                End If
            End With
        Next shp
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

        Dim shp As Excel.Shape

        For Each shp In appInstance.ActiveSheet.Shapes
            With shp
                If shp.AutoShapeType = MsoAutoShapeType.msoShapeRoundedRectangle Or _
                    shp.AutoShapeType = MsoAutoShapeType.msoShapeIsoscelesTriangle Or _
                    shp.AutoShapeType = MsoAutoShapeType.msoShapeMixed Then
                    .Delete()
                End If
            End With
        Next shp

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

        ' jetzt wird die Zuordnung Projektname und Shapge ID gelöscht ... 
        ShowProjekte.shpListe.Clear()

        appInstance.EnableEvents = formerEE

    End Sub

    Public Sub ClearPlanTafelfromOptArrows()

        Dim formerEE As Boolean = appInstance.EnableEvents
        appInstance.EnableEvents = False



        ' jetzt müssen alle Shapes, die Optmmierungs-Arrows sind, gelöscht werden ....

        Dim shp As Excel.Shape

        For Each shp In appInstance.ActiveSheet.Shapes
            With shp
                If shp.AutoShapeType = MsoAutoShapeType.msoShapeRightArrow Or _
                    shp.AutoShapeType = MsoAutoShapeType.msoShapeLeftArrow Then
                    .Delete()
                End If
            End With
        Next shp



        appInstance.EnableEvents = formerEE


    End Sub
    ''' <summary>
    ''' zeichnet die Plantafel mit den Projekten neu; 
    ''' versucht dabei immer die alte Position der Projekte zu übernehmen 
    ''' </summary>
    ''' <remarks></remarks>
    Public Sub awinZeichnePlanTafel()


        Dim zeile As Integer
        Dim nrCols As Integer




        '
        ' wieviele Spalten sind im Vorlagen Sheet relevant 
        '
        nrCols = 4
        zeile = 2

        Dim pname As String
        Dim tryzeile As Integer

        For Each kvp As KeyValuePair(Of String, clsProjekt) In ShowProjekte.Liste
            pname = kvp.Key
            tryzeile = kvp.Value.tfZeile
            If tryzeile <= 1 Then
                tryzeile = -1
            End If
            Call ZeichneProjektinPlanTafel(pname, tryzeile, False) ' es wird versucht, an der alten Stelle zu zeichnen 
        Next



    End Sub
    ''' <summary>
    ''' zeichnet die Plantafel mit den Projekten neu; 
    ''' beachtet keine vorherige Position der Projekte
    ''' </summary>
    ''' <remarks></remarks>
    Public Sub awinZeichnePlanTafelFromScratch()

        ' die Projekte werden in folgender Reihenfolge gezeichnet:
        ' sortiert nach : 
        ' 1. Länge (absteigend)
        ' 2. Start Monat (aufsteigend)
        ' 3. Vorlagen-Typ (aufsteigend)
        ' 4. Name (aufsteigend) 


        Dim letztezeile As Integer, zeile As Integer
        Dim nrCols As Integer
        Dim rng As Range



        '
        ' wieviele Spalten sind im Vorlagen Sheet relevant 
        '
        nrCols = 4
        zeile = 2



        With appInstance.Worksheets(arrWsNames(2))
            ' jetzt werden die Projekte in die Tabelle Vorlage für das nachfolgende Sortieren geschrieben ... 
            For Each kvp As KeyValuePair(Of String, clsProjekt) In ShowProjekte.Liste

                .cells(zeile, 1).value = kvp.Value.name
                .cells(zeile, 2).value = kvp.Value.VorlagenName
                .cells(zeile, 3).value = kvp.Value.Start + kvp.Value.StartOffset
                .cells(zeile, 4).value = kvp.Value.Dauer

                zeile = zeile + 1
            Next kvp


            letztezeile = zeile - 1
            rng = .Range(.Cells(2, 1), .Cells(letztezeile, nrCols))


            ' Sortieren - erst nach Dauer, dann nach Start, dann nach Vorlagen Name, dann nach Name
            With .Sort
                ' Bestehende Sortierebenen löschen
                .SortFields.Clear()
                ' Neue Sortierebenen hinzufügen
                .SortFields.Add(Key:=appInstance.Range("D2:D" & letztezeile), Order:=XlSortOrder.xlDescending)
                .SortFields.Add(Key:=appInstance.Range("C2:C" & letztezeile), Order:=XlSortOrder.xlAscending)
                .SortFields.Add(Key:=appInstance.Range("B2:B" & letztezeile), Order:=XlSortOrder.xlAscending)
                .SortFields.Add(Key:=appInstance.Range("A2:A" & letztezeile), Order:=XlSortOrder.xlAscending)
                .SetRange(rng)
                .Apply()
            End With

            zeile = 2
            Dim pname As String
            While zeile <= letztezeile
                pname = .cells(zeile, 1).value
                Call ZeichneProjektinPlanTafel(pname, -1, False) ' es muss nicht an der alten Stelle gezeichnet werden ... 
                zeile = zeile + 1
            End While

            ' die Werte wieder löschen ...
            rng.Clear()
        End With

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
                    referenceValue = .getDeviationfromAverage(myCollection, avgValue, False)
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
                                        currentValue = .getDeviationfromAverage(myCollection, avgValue, False)
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
                            currentValue = .getDeviationfromAverage(myCollection, avgValue, False)
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

    Public Sub awinCalculateOptimization1(ByVal diagrammTyp As String, ByRef myCollection As Collection, _
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


            If diagrammTyp = DiagrammTypen(1) Or diagrammTyp = DiagrammTypen(2) Or diagrammTyp = DiagrammTypen(4) Then

                ' to do Liste aufbauen
                For Each kvp As KeyValuePair(Of String, clsProjekt) In ShowProjekte.Liste
                    ' testweise
                    kvp.Value.earliestStart = -3
                    kvp.Value.latestStart = 3
                    If relevantForOptimization(kvp.Value) Then
                        toDoListe.Add(kvp.Key, kvp.Key)
                    End If
                Next kvp

                bestValue = berechneOptimierungsWert(diagrammTyp, myCollection)
                lokalesOptimum.bestValue = bestValue
                lokalesOptimum.projectName = " "
                OptimierungsErgebnis.Clear()

                NrLoops = 0
                NrArgExceptions = 0


                'While newReferenceValue < referenceValue And toDoListe.Count > 0
                Dim Abbruch As Boolean = False
                While toDoListe.Count > 0 And Not Abbruch

                    Dim i As Integer
                    Dim kvp As clsProjekt

                    For i = 1 To toDoListe.Count
                        kvp = ShowProjekte.getProject(toDoListe.Item(i))

                        startoffset = 0

                        ' hier wird der beste Wert für das einzelne Projekt gesucht ....  
                        'For versatz = kvp.earliestStart To kvp.latestStart
                        For versatz = -3 To 3
                            If versatz <> 0 Then
                                kvp.StartOffset = versatz
                                currentValue = berechneOptimierungsWert(diagrammTyp, myCollection)

                                If currentValue < bestValue Then
                                    bestValue = currentValue
                                    startoffset = versatz
                                End If
                            End If
                        Next versatz

                        ' zurücksetzen des StartOffsets im Projekt, weil hier ja erst verschiedene Konstellationen probiert werden  
                        kvp.StartOffset = 0

                        If startoffset <> 0 Then ' es gab eine Verbesserung 
                            'lokalesOptimum = New clsOptimizationObject
                            With lokalesOptimum
                                If bestValue < .bestValue Then
                                    .projectName = kvp.name
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

    Public Function berechneOptimierungsWert(ByRef DiagrammTyp As String, ByRef myCollection As Collection) As Double
        Dim value As Double
        Dim avgValue As Double

        If DiagrammTyp = DiagrammTypen(1) Or DiagrammTyp = DiagrammTypen(4) Then
            value = ShowProjekte.getbadCostOfRole(myCollection)
        ElseIf DiagrammTyp = DiagrammTypen(2) Then
            avgValue = ShowProjekte.getAverage(myCollection, False)
            value = ShowProjekte.getDeviationfromAverage(myCollection, avgValue, False)
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
                bereichsEnde = .Start + .Dauer - 1 + .latestStart

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

    Public Sub awinZeichnePhasen(ByVal nameList As Collection, ByVal farbTyp As Integer, ByVal numberIt As Boolean)

        'Dim request As New Request(awinSettings.databaseName)
        Dim singleShp As Excel.Shape
        Dim hproj As New clsProjekt
        Dim vglName As String = " "
        Dim pName As String
        Dim ok As Boolean = True

        Dim awinSelection As Excel.ShapeRange

        Dim formerEE As Boolean = appInstance.EnableEvents
        Dim formerSU As Boolean = appInstance.ScreenUpdating
        appInstance.EnableEvents = False
        appInstance.ScreenUpdating = False

        enableOnUpdate = False

        Try
            awinSelection = appInstance.ActiveWindow.Selection.ShapeRange
        Catch ex As Exception
            awinSelection = Nothing
        End Try

        If Not awinSelection Is Nothing Then

            ' jetzt die Aktion durchführen ...

            For Each singleShp In awinSelection
                ok = True
                With singleShp
                    If .AutoShapeType = MsoAutoShapeType.msoShapeRoundedRectangle Or _
                        .AutoShapeType = MsoAutoShapeType.msoShapeMixed Then

                        Try
                            hproj = ShowProjekte.getProject(singleShp.Name)
                        Catch ex As Exception
                            ok = False
                        End Try

                        If ok Then

                            Try
                                pName = hproj.name
                                Call zeichnePhasenInProjekt(hproj, nameList, farbTyp, False, 0, False)

                            Catch ex As Exception

                            End Try


                        End If

                    End If
                End With
            Next

        Else
            ' tue es für alle Projekte in Showprojekte 


            Dim todoListe As New SortedList(Of Long, clsProjekt)
            Dim key As Long

            For Each kvp As KeyValuePair(Of String, clsProjekt) In ShowProjekte.Liste

                key = 10000 * kvp.Value.tfZeile + kvp.Value.tfspalte
                todoListe.Add(key, kvp.Value)

            Next
            Dim msNumber As Integer = 1

            For Each kvp As KeyValuePair(Of Long, clsProjekt) In todoListe

                ' wenn ein Zeitraum gesetzt ist, dann nur anzeigen, was in diesem Zeitraum liegt 
                If showRangeLeft < showRangeRight And showRangeLeft > 0 Then
                    Call zeichnePhasenInProjekt(kvp.Value, nameList, farbTyp, showRangeLeft, showRangeRight, False, 0, False)
                Else
                    ' von jedem Projekt die Phasen anzeigen 
                    Call zeichnePhasenInProjekt(kvp.Value, nameList, farbTyp, False, 0, False)
                End If


            Next





        End If

        Call awinDeSelect()

        enableOnUpdate = True
        appInstance.EnableEvents = formerEE
        appInstance.ScreenUpdating = formerSU



    End Sub



    ''' <summary>
    ''' zeichnet für interaktiven wie Report Modus die Milestones 
    ''' 0: grau, 1: grün, 2: gelb, 3:rot, 4: alle
    ''' </summary>
    ''' <param name="farbTyp">welcher Typus soll gezeichnet werden </param>
    ''' <remarks></remarks>
    Public Sub awinZeichneMilestones(ByVal nameList As SortedList(Of String, String), ByVal farbTyp As Integer, ByVal numberIt As Boolean)

        'Dim request As New Request(awinSettings.databaseName)
        Dim singleShp As Excel.Shape
        Dim hproj As New clsProjekt
        Dim vglName As String = " "
        Dim pName As String
        Dim ok As Boolean = True

        Dim awinSelection As Excel.ShapeRange

        Dim formerEE As Boolean = appInstance.EnableEvents
        Dim formerSU As Boolean = appInstance.ScreenUpdating
        appInstance.EnableEvents = False
        appInstance.ScreenUpdating = False

        enableOnUpdate = False

        Try
            awinSelection = appInstance.ActiveWindow.Selection.ShapeRange
        Catch ex As Exception
            awinSelection = Nothing
        End Try

        If Not awinSelection Is Nothing Then

            ' jetzt die Aktion durchführen ...

            For Each singleShp In awinSelection
                ok = True
                With singleShp
                    If .AutoShapeType = MsoAutoShapeType.msoShapeRoundedRectangle Or _
                        .AutoShapeType = MsoAutoShapeType.msoShapeMixed Then

                        Try
                            hproj = ShowProjekte.getProject(singleShp.Name)
                        Catch ex As Exception
                            ok = False
                        End Try

                        If ok Then

                            Try
                                pName = hproj.name
                                Call zeichneResultMilestonesInProjekt(hproj, nameList, farbTyp, False, False, 0, False)
                            Catch ex As Exception

                            End Try


                        End If

                    End If
                End With
            Next

        Else
            ' tue es für alle Projekte in Showprojekte 


            Dim todoListe As New SortedList(Of Long, clsProjekt)
            Dim key As Long

            For Each kvp As KeyValuePair(Of String, clsProjekt) In ShowProjekte.Liste

                key = 10000 * kvp.Value.tfZeile + kvp.Value.tfspalte
                todoListe.Add(key, kvp.Value)

            Next
            Dim msNumber As Integer = 1

            For Each kvp As KeyValuePair(Of Long, clsProjekt) In todoListe

                Call zeichneResultMilestonesInProjekt(kvp.Value, nameList, farbTyp, True, numberIt, msNumber, False)

            Next





        End If

        Call awinDeSelect()

        enableOnUpdate = True
        appInstance.EnableEvents = formerEE
        appInstance.ScreenUpdating = formerSU



    End Sub

    ''' <summary>
    ''' zeichnet das Projekt "pname" in die Plantafel; 
    ''' wenn es bereits vorhanden ist: keine Aktion  
    ''' </summary>
    ''' <param name="pname"></param>
    ''' <remarks></remarks>
    Public Sub ZeichneProjektinPlanTafel(ByVal pname As String, ByVal tryzeile As Integer, ByVal showresults As Boolean)


        Dim drawphases As Boolean = My.Settings.drawPhases
        Dim phasenName As String
        Dim phaseShapeName As String

        Dim start As Integer
        Dim laenge As Integer
        Dim status As String
        Dim pMarge As Double
        Dim pcolor As Object, schriftfarbe As Object
        Dim schriftgroesse As Integer
        Dim zeile As Integer
        Dim hproj As clsProjekt
        Dim top As Double, left As Double, width As Double, height As Double
        Dim shpElement As Excel.Shape, groupShpElement As Excel.Shape
        'Dim shpExistsAlready As Boolean
        Dim shpUID As String
        'Dim tmpshapes As Excel.Shapes = appInstance.ActiveSheet.shapes
        Dim worksheetShapes As Excel.Shapes
        Dim heute As Date = Date.Now

        Dim shpExists As Boolean


        Try

            worksheetShapes = CType(appInstance.Worksheets(arrWsNames(3)), Excel.Worksheet).Shapes

        Catch ex As Exception
            Throw New Exception("in ZeichneProjektinPlanTafel : keine Shapes Zuordnung möglich ")
        End Try

        Try
            hproj = ShowProjekte.getProject(pname)
            With hproj
                laenge = .Dauer
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
                shpElement = worksheetShapes.Item(pname)
                shpExists = True
            Catch ex As Exception
                shpExists = False
                shpElement = Nothing
            End Try
        Else
            shpExists = False
            shpElement = Nothing
        End If



        '
        ' ist dort überhaupt Platz ? wenn nicht, dann Zeile mit freiem Platz suchen ...
        If tryzeile < 2 Then
            tryzeile = 2
        End If

        Dim myCollection As New Collection
        myCollection.Add(pname)
        zeile = findeMagicBoardPosition(myCollection, pname, tryzeile, start, laenge)


        Dim formerEE As Boolean = appInstance.EnableEvents
        appInstance.EnableEvents = False



        If shpExists Then
            
            If drawphases Then

                For i = 1 To hproj.CountPhases
                    phasenName = hproj.getPhase(i).name
                    phaseShapeName = pname & "#" & phasenName & "#" & i.ToString

                    Try
                        shpElement = worksheetShapes.Item(phaseShapeName)
                        Call defineShapeAppearance(hproj, shpElement, i)
                    Catch ex As Exception

                    End Try


                Next

            Else
                Call defineShapeAppearance(hproj, shpElement)
            End If


            ' jetzt prüfen, ob die Ergebnis Meilensteine angezeigt werden sollen oder auch schon angezeigt werden  
            'If showresults Then

            'Call zeichneResultMilestonesInPlantafel(hproj)

            'End If

        Else

            ' ///////////////
            ' Start neuer code 
            ' ///////////////

            ' hier wird der vorher bestimmte Wert gesetzt, wo das Shape gezeichnet werden kann 
            hproj.tfZeile = zeile

            If drawphases And hproj.CountPhases > 1 Then

                Dim shapeGroupListe() As Object
                Dim anzGroupElemente As Integer = 0
                Dim shapesCollection As New Collection


                shpElement = Nothing

                For i = 1 To hproj.CountPhases
                    phasenName = hproj.getPhase(i).name
                    hproj.CalculateShapeCoord(i, top, left, width, height)



                    Try
                        If i = 1 Then
                            shpElement = worksheetShapes.AddShape(Type:=Microsoft.Office.Core.MsoAutoShapeType.msoShapePentagon, _
                               Left:=left, Top:=top, Width:=width, Height:=height)
                        Else
                            shpElement = worksheetShapes.AddShape(Type:=Microsoft.Office.Core.MsoAutoShapeType.msoShapeChevron, _
                               Left:=left, Top:=top, Width:=width, Height:=height)
                        End If
                    Catch ex As Exception
                        Throw New Exception("in zeichneShapeOfProject : keine Shape-Erstellung möglich ...  ")
                    End Try


                    phaseShapeName = pname & "#" & phasenName & "#" & i.ToString
                    With shpElement
                        .Name = phaseShapeName
                        .Title = phasenName
                    End With

                    Call defineShapeAppearance(hproj, shpElement, i)


                    Try
                        shapesCollection.Add(phaseShapeName, Key:=phaseShapeName)
                    Catch ex As Exception

                    End Try


                Next

                ' hier jetzt noch den Text ergänzen

                ' hier werden die Shapes gruppiert
                anzGroupElemente = shapesCollection.Count

                If anzGroupElemente > 1 Then
                    ' es macht nur Sinn zu gruppieren, wenn es mehr als 1 Element ist ....

                    ReDim shapeGroupListe(anzGroupElemente - 1)
                    For i = 1 To anzGroupElemente
                        shapeGroupListe(i - 1) = shapesCollection.Item(i)
                    Next

                    Dim ShapeGroup As Excel.ShapeRange
                    ShapeGroup = worksheetShapes.Range(shapeGroupListe)
                    groupShpElement = ShapeGroup.Group()

                Else
                    ' in diesem Fall besteht das Projekt nur aus einer einzigen Phase
                    groupShpElement = shpElement

                End If

                Try
                    With groupShpElement
                        .Name = pname
                        hproj.shpUID = .ID.ToString
                        hproj.tfZeile = zeile
                    End With
                Catch ex As Exception
                    Throw New Exception("in zeichneShapeOfProject : dem shape kann kein Name zugewiesen werden ....   ")
                End Try

                ' jetzt muss das neue Shape in der ShowProjekte.ShapeListe eingetragen werden ..
                ShowProjekte.AddShape(pname, shpUID:=groupShpElement.ID.ToString)



            Else

                With hproj
                    .CalculateShapeCoord(top, left, width, height)
                    .tfZeile = zeile
                End With

                shpElement = worksheetShapes.AddShape(Type:=Microsoft.Office.Core.MsoAutoShapeType.msoShapeRoundedRectangle, _
                        Left:=left, Top:=top, Width:=width, Height:=height)


                With shpElement
                    .Name = pname
                    hproj.shpUID = .ID.ToString
                End With

                Call defineShapeAppearance(hproj, shpElement)

                ' jetzt muss das neue Shape in der ShowProjekte.ShapeListe eingetragen werden ..
                ShowProjekte.AddShape(pname, shpUID:=shpElement.ID.ToString)
            End If

            '' //////////////
            '' Start alter Code 
            '' //////////////
            'With hproj
            '    '.tfSpalte = start
            '    .tfZeile = zeile
            '    .CalculateShapeCoord(top, left, width, height)
            'End With

            'shpElement = .Shapes.AddShape(Type:=Microsoft.Office.Core.MsoAutoShapeType.msoShapeRoundedRectangle, _
            '            left:=left, top:=top, width:=width, height:=height)

            'With shpElement
            '    .Name = pname
            '    hproj.shpUID = .ID.ToString
            'End With

            'Call defineShapeAppearance(hproj, shpElement)


            '' jetzt muss das neue Shape in der ShowProjekte.ShapeListe eingetragen werden ..
            'ShowProjekte.AddShape(pname, shpUID:=shpElement.ID.ToString)

        End If


        If roentgenBlick.isOn Then
            With roentgenBlick
                Call awinShowNeedsofProject1(mycollection:=.myCollection, type:=.type, projektname:=pname)
            End With
        End If

        appInstance.EnableEvents = formerEE

    End Sub

    Public Sub zeichneStatusSymbolInPlantafel(ByVal hproj As clsProjekt, ByVal number As Integer)
        Dim top As Double, left As Double, height As Double, width As Double
        Dim worksheetShapes As Excel.Shapes
        Dim shpElement As Excel.Shape
        Dim resultShape As Excel.Shape
        Dim shpName As String
        Dim heute As Date = Date.Now

        With CType(appInstance.Worksheets(arrWsNames(3)), Excel.Worksheet)
            worksheetShapes = .Shapes


            shpName = hproj.name & "#S" & "Status"
            ' existiert das schon ? 
            Try
                shpElement = worksheetShapes.Item(shpName)
            Catch ex As Exception
                shpElement = Nothing
            End Try

            If shpElement Is Nothing Then

                hproj.calculateStatusCoord(heute, top, left, width, height)
                resultShape = .Shapes.AddShape(Type:=Microsoft.Office.Core.MsoAutoShapeType.msoShapeOval, _
                                                Left:=left, Top:=top, Width:=width, Height:=height)

                With resultShape
                    .Name = shpName
                    .Title = "Status"
                    .AlternativeText = ""
                End With

                Call defineStatusAppearance(hproj, number, resultShape)
                'shapesCollection.Add(resultShape.Name)

            End If


        End With
    End Sub

    Public Sub zeichneMilestones(ByVal nameList As SortedList(Of String, String), ByVal farbTyp As Integer, ByVal numberIt As Boolean)
        ' tue es für alle Projekte in Showprojekte 


        Dim todoListe As New SortedList(Of Long, clsProjekt)
        Dim key As Long
        Dim formerEE As Boolean = appInstance.EnableEvents
        Dim formereO As Boolean = enableOnUpdate

        appInstance.EnableEvents = False
        enableOnUpdate = False

        For Each kvp As KeyValuePair(Of String, clsProjekt) In ShowProjekte.Liste

            key = 10000 * kvp.Value.tfZeile + kvp.Value.tfspalte
            todoListe.Add(key, kvp.Value)

        Next

        Dim msNumber As Integer = 1

        For Each kvp As KeyValuePair(Of Long, clsProjekt) In todoListe

            Call zeichneResultMilestonesInProjekt(kvp.Value, nameList, farbTyp, True, numberIt, msNumber, False)

        Next

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
    ''' <param name="showOnlyWithinTimeFrame">
    ''' gibt an , ob der aktuell betrachtete Zeitraum berücksichtigt werden soll</param>
    ''' <param name="numberIt">
    ''' gibt an, ob der Meilenstein nummeriert werden soll</param>
    ''' <param name="msNumber">
    ''' gibt die Nummer an, aber nummeriert werden soll</param>
    ''' <param name="report">
    ''' gibt an, ob vom Reporting aufgerufen
    ''' </param>
    ''' <remarks></remarks>
    Public Sub zeichneResultMilestonesInProjekt(ByVal hproj As clsProjekt, ByVal namenListe As SortedList(Of String, String), ByVal farbTyp As Integer, ByVal showOnlyWithinTimeFrame As Boolean, _
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

        ' sollen nur die in der Namenliste aufgeführten Meilensteine gezeichnet werden ? 
        Try
            If namenListe.Count > 0 Then
                onlyFew = True
            Else
                onlyFew = False
            End If
        Catch ex As Exception
            onlyFew = False
        End Try

        With appInstance.Worksheets(arrWsNames(3))

            worksheetShapes = .shapes


            For p = 1 To hproj.CountPhases

                Dim cphase As clsPhase = hproj.getPhase(p)


                For r = 1 To cphase.CountResults
                    Dim cResult As clsResult
                    Dim cBewertung As clsBewertung
                    Dim nameIstInListe As Boolean

                    cResult = cphase.getResult(r)

                    If namenListe.ContainsKey(cResult.name) Then
                        nameIstInListe = True
                    Else
                        nameIstInListe = False
                    End If

                    cBewertung = cResult.getBewertung(1)


                    resultColumn = getColumnOfDate(cResult.getDate)

                    If farbTyp = 4 Or farbTyp = cBewertung.colorIndex Then
                        ' es muss nur etwas gemacht werden , wenn entweder alle Farben gezeichnet werden oder eben die übergebene

                        If (showOnlyWithinTimeFrame And (resultColumn < showRangeLeft Or resultColumn > showRangeRight)) Or _
                            (onlyFew And Not nameIstInListe) Then
                            ' nichts machen 
                        Else
                            hproj.calculateResultCoord(cResult.getDate, top, left, width, height)

                            If DateDiff(DateInterval.Month, cResult.getDate, heute) >= 0 Then
                                ' zeichne Raute für vergangenes Ziel 


                                shpName = hproj.name & "#" & cphase.name & "#M" & r.ToString
                                ' existiert das schon ? 
                                Try
                                    shpElement = worksheetShapes.Item(shpName)
                                Catch ex As Exception
                                    shpElement = Nothing
                                End Try

                                If shpElement Is Nothing Then

                                    If report Then
                                        top = top - 0.7 * boxWidth
                                    End If
                                    resultShape = .Shapes.AddShape(Type:=Microsoft.Office.Core.MsoAutoShapeType.msoShapeDiamond, _
                                                                    left:=left, top:=top, width:=width, height:=height)

                                    With resultShape
                                        .Name = shpName
                                        .Title = cResult.name
                                        .AlternativeText = cBewertung.description
                                    End With

                                    If numberIt Then
                                        Call defineResultAppearance(hproj, msNumber, resultShape, cBewertung)
                                        msNumber = msNumber + 1
                                    Else
                                        Call defineResultAppearance(hproj, 0, resultShape, cBewertung)
                                    End If


                                    'shapesCollection.Add(resultShape.Name)

                                End If


                            Else
                                ' zeichne Kreis für zukünftiges Ziel ; Kreis soll etwas kleiner sein als die Raute ...


                                shpName = hproj.name & "#" & cphase.name & "#M" & r.ToString
                                ' existiert das schon ? 
                                Try
                                    shpElement = worksheetShapes.Item(shpName)
                                Catch ex As Exception
                                    shpElement = Nothing
                                End Try

                                If shpElement Is Nothing Then

                                    'resultShape = .Shapes.AddShape(Type:=Microsoft.Office.Core.MsoAutoShapeType.msoShapeOval, _
                                    '                                left:=left + 0.05 * width, top:=top + 0.05 * height, width:=width * 0.9, height:=height * 0.9)

                                    If report Then
                                        top = top - 0.7 * boxWidth
                                    End If
                                    resultShape = .Shapes.AddShape(Type:=Microsoft.Office.Core.MsoAutoShapeType.msoShapeDiamond, _
                                                                    left:=left, top:=top, width:=width, height:=height)
                                    With resultShape
                                        .Name = shpName
                                        .Title = cResult.name
                                        .AlternativeText = cBewertung.description
                                    End With

                                    If numberIt Then
                                        Call defineResultAppearance(hproj, msNumber, resultShape, cBewertung)
                                        msNumber = msNumber + 1
                                    Else
                                        Call defineResultAppearance(hproj, 0, resultShape, cBewertung)
                                    End If

                                    'shapesCollection.Add(resultShape.Name)

                                End If

                            End If
                        End If


                    End If



                Next

            Next

            ' hier werden die Shapes gruppiert
            'anzGroupElemente = shapesCollection.Count

            'If anzGroupElemente > 1 Then
            '    ' es macht nur Sinn zu gruppieren, wenn es mehr als 1 Element ist ....

            '    ReDim shapeGroupListe(anzGroupElemente - 1)
            '    For i = 1 To anzGroupElemente
            '        shapeGroupListe(i - 1) = shapesCollection.Item(i)
            '    Next

            '    Dim ShapeGroup As Excel.ShapeRange
            '    ShapeGroup = worksheetShapes.Range(shapeGroupListe)
            '    shpElement = ShapeGroup.Group()


            '    With shpElement
            '        .Name = key
            '    End With

            '    'ShowProjekte.AddShape(pname, shpUID:=shpElement.ID.ToString)
            'Else
            '    ' nichts zu tun
            'End If




        End With


    End Sub

    Public Sub zeichnePhasenInProjekt(ByVal hproj As clsProjekt, ByVal namenListe As Collection, ByVal farbTyp As Integer, _
                                                      ByVal numberIt As Boolean, ByRef msNumber As Integer, ByVal report As Boolean)

        Dim top1 As Double, left1 As Double, top2 As Double, left2 As Double
        Dim nummer As Integer, gesamtZahl As Integer
        Dim phasenShape As Excel.Shape
        Dim worksheetShapes As Excel.Shapes
        Dim heute As Date = Date.Now
        Dim alreadyGroup As Boolean = False
        Dim shpElement As Excel.Shape
        Dim shpName As String

        Dim onlyFew As Boolean
        Dim nameIstInListe As Boolean
        Dim linienDicke As Double = 2.0


        ' als wievielte Phase wird das Shape gezeichnet ... 
        nummer = 1

        ' sollen nur die in der Namenliste aufgeführten Phasen gezeichnet werden ? 
        Try
            If namenListe.Count > 0 Then
                onlyFew = True
                gesamtZahl = namenListe.Count
            Else
                onlyFew = False
                gesamtZahl = hproj.CountPhases
            End If
        Catch ex As Exception
            onlyFew = False
        End Try


        With appInstance.Worksheets(arrWsNames(3))

            worksheetShapes = .shapes


            For p = 1 To hproj.CountPhases

                Dim cphase As clsPhase = hproj.getPhase(p)

                Try
                    nameIstInListe = namenListe.Contains(cphase.name)
                Catch ex As Exception
                    nameIstInListe = False
                End Try



                If onlyFew And Not nameIstInListe Then
                    ' nichts machen 
                Else

                    linienDicke = boxHeight * 0.3

                    ' wieder gültig machen, wenn die Liniendicke in Abhängigkeit sein soll von der Anzahl der gezeigten Phasen ...
                    'If gesamtZahl > 1 Then
                    '    If gesamtZahl <= 10 Then
                    '        linienDicke = linienDicke * (1 - 0.08 * gesamtZahl)
                    '    Else
                    '        linienDicke = linienDicke * 0.2
                    '    End If
                    'End If


                    Try
                        cphase.calculateLineCoord(hproj.tfZeile, nummer, gesamtZahl, top1, left1, top2, left2, linienDicke)
                    Catch ex As Exception
                        Throw New ArgumentException(ex.Message)
                    End Try

                    nummer = nummer + 1

                    shpName = hproj.name & "#" & cphase.name
                    ' existiert das schon ? 
                    Try
                        shpElement = worksheetShapes.Item(shpName)
                    Catch ex As Exception
                        shpElement = Nothing
                    End Try

                    If shpElement Is Nothing Then


                        phasenShape = .Shapes.AddConnector(MsoConnectorType.msoConnectorStraight, left1, top1, left2, top2)

                        With phasenShape
                            .Name = shpName
                            .Title = cphase.name
                            .AlternativeText = ""
                        End With

                        If numberIt Then
                            Call defineLineAppearance(hproj, cphase, msNumber, phasenShape, linienDicke)
                            msNumber = msNumber + 1
                        Else
                            Call defineLineAppearance(hproj, cphase, 0, phasenShape, linienDicke)
                        End If


                    End If




                End If

            Next

            ' hier werden die Shapes gruppiert
            'anzGroupElemente = shapesCollection.Count

            'If anzGroupElemente > 1 Then
            '    ' es macht nur Sinn zu gruppieren, wenn es mehr als 1 Element ist ....

            '    ReDim shapeGroupListe(anzGroupElemente - 1)
            '    For i = 1 To anzGroupElemente
            '        shapeGroupListe(i - 1) = shapesCollection.Item(i)
            '    Next

            '    Dim ShapeGroup As Excel.ShapeRange
            '    ShapeGroup = worksheetShapes.Range(shapeGroupListe)
            '    shpElement = ShapeGroup.Group()


            '    With shpElement
            '        .Name = key
            '    End With

            '    'ShowProjekte.AddShape(pname, shpUID:=shpElement.ID.ToString)
            'Else
            '    ' nichts zu tun
            'End If




        End With


    End Sub


    Public Sub zeichnePhasenInProjekt(ByVal hproj As clsProjekt, ByVal namenListe As Collection, ByVal farbTyp As Integer, ByVal vonMonth As Integer, ByVal bisMonth As Integer, _
                                                          ByVal numberIt As Boolean, ByRef msNumber As Integer, ByVal report As Boolean)

        Dim top1 As Double, left1 As Double, top2 As Double, left2 As Double
        Dim nummer As Integer, gesamtZahl As Integer
        Dim phasenShape As Excel.Shape
        Dim worksheetShapes As Excel.Shapes
        Dim heute As Date = Date.Now
        Dim alreadyGroup As Boolean = False
        Dim shpElement As Excel.Shape
        Dim shpName As String
        Dim todoListe As New Collection

        Dim onlyFew As Boolean
        Dim nameIstInListe As Boolean
        Dim linienDicke As Double = 2.0
        Dim ok As Boolean = True


        ' als wievielte Phase wird das Shape gezeichnet ... 
        nummer = 1

        ' sollen nur die in der Namenliste aufgeführten Phasen gezeichnet werden ? 
        Try
            If namenListe.Count > 0 Then
                onlyFew = True
                gesamtZahl = namenListe.Count
            Else
                onlyFew = False
                gesamtZahl = hproj.CountPhases
            End If
        Catch ex As Exception
            onlyFew = False
        End Try


        Try
            'bringt eine List von Phasen Namen zurück, die den angegeben Zeitraum berühren / überdecken
            todoListe = hproj.withinTimeFrame(0, vonMonth, bisMonth)

        Catch ex As Exception

        End Try

        With appInstance.Worksheets(arrWsNames(3))

            worksheetShapes = .shapes

            Dim cphase As clsPhase

            ' in der todoListe stehen jetzt nur Phasen, die den angegeben Zeitraum betreffen 
            For p = 1 To todoListe.Count

                Dim phaseName As String = todoListe(p)
                cphase = hproj.getPhase(phaseName)

                Try
                    ' soll diese Phase überhaupt gezeigt werden ? 
                    nameIstInListe = namenListe.Contains(phaseName)

                Catch ex As Exception
                    nameIstInListe = False
                End Try



                If onlyFew And Not nameIstInListe Then
                    ' nichts machen 
                Else

                    linienDicke = boxHeight * 0.3
                    'If gesamtZahl > 1 Then
                    '    If gesamtZahl <= 10 Then
                    '        linienDicke = linienDicke * (1 - 0.08 * gesamtZahl)
                    '    Else
                    '        linienDicke = linienDicke * 0.2
                    '    End If
                    'End If

                    Try
                        cphase.calculateLineCoord(hproj.tfZeile, nummer, gesamtZahl, top1, left1, top2, left2, linienDicke)
                    Catch ex As Exception
                        ok = False
                    End Try



                    If ok Then
                        nummer = nummer + 1

                        shpName = hproj.name & "#" & cphase.name
                        ' existiert das schon ? 
                        Try
                            shpElement = worksheetShapes.Item(shpName)
                        Catch ex As Exception
                            shpElement = Nothing
                        End Try

                        If shpElement Is Nothing Then


                            phasenShape = .Shapes.AddConnector(MsoConnectorType.msoConnectorStraight, left1, top1, left2, top2)

                            With phasenShape
                                .Name = shpName
                                .Title = cphase.name
                                .AlternativeText = ""
                            End With

                            If numberIt Then
                                Call defineLineAppearance(hproj, cphase, msNumber, phasenShape, linienDicke)
                                msNumber = msNumber + 1
                            Else
                                Call defineLineAppearance(hproj, cphase, 0, phasenShape, linienDicke)
                            End If


                        End If

                    End If






                End If

                ok = True
            Next

            ' hier werden die Shapes gruppiert
            'anzGroupElemente = shapesCollection.Count

            'If anzGroupElemente > 1 Then
            '    ' es macht nur Sinn zu gruppieren, wenn es mehr als 1 Element ist ....

            '    ReDim shapeGroupListe(anzGroupElemente - 1)
            '    For i = 1 To anzGroupElemente
            '        shapeGroupListe(i - 1) = shapesCollection.Item(i)
            '    Next

            '    Dim ShapeGroup As Excel.ShapeRange
            '    ShapeGroup = worksheetShapes.Range(shapeGroupListe)
            '    shpElement = ShapeGroup.Group()


            '    With shpElement
            '        .Name = key
            '    End With

            '    'ShowProjekte.AddShape(pname, shpUID:=shpElement.ID.ToString)
            'Else
            '    ' nichts zu tun
            'End If




        End With


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
                .ForeColor.RGB = pColor
                .Transparency = 0
            End With

            With .Fill
                '.Visible = msoTrue
                .ForeColor.RGB = pColor
                .ForeColor.TintAndShade = 0
                .ForeColor.Brightness = -0.25

                If roentgenBlick.isOn Then
                    .Transparency = 0.8
                Else
                    .Transparency = 0.0
                End If

                .Solid()

            End With


            .TextFrame2.TextRange.Text = ""
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







        End With


    End Sub

    Public Sub defineLineAppearance(ByVal myproject As clsProjekt, ByVal myphase As clsPhase, ByVal lnumber As Integer, ByRef myShape As Excel.Shape, ByVal linienDicke As Double)
        Dim pColor As Long

        With myphase

            pColor = .Farbe

        End With

        With myShape

            With .Line
                .Visible = Microsoft.Office.Core.MsoTriState.msoTrue
                .ForeColor.RGB = pColor
                .Transparency = 0
                .Weight = linienDicke
            End With


            '.TextFrame2.TextRange.Text = ""
            'If lnumber > 0 And Not roentgenBlick.isOn Then

            '    With .TextFrame2
            '        .MarginLeft = 0
            '        .MarginRight = 0
            '        .MarginBottom = 0
            '        .MarginTop = 0
            '        .WordWrap = Microsoft.Office.Core.MsoTriState.msoFalse
            '        .VerticalAnchor = MsoVerticalAnchor.msoAnchorMiddle
            '        .HorizontalAnchor = MsoHorizontalAnchor.msoAnchorCenter
            '        .TextRange.Text = lnumber.ToString
            '        .TextRange.Font.Size = awinSettings.fontsizeLegend
            '        .TextRange.Font.Fill.ForeColor.RGB = RGB(255, 255, 255)
            '    End With


            'End If


        End With


    End Sub


    Public Sub defineResultAppearance(ByVal myproject As clsProjekt, ByVal number As Integer, ByRef resultShape As Excel.Shape, ByVal bewertung As clsBewertung)
        Dim pcolor As Object
        Dim status As String

        With myproject
            pcolor = .farbe
            status = .Status
        End With


        With resultShape

            With .Line
                '.Visible = Microsoft.Office.Core.MsoTriState.msoTrue
                .Visible = MsoTriState.msoTrue
                .ForeColor.RGB = RGB(255, 255, 255)
                '.ForeColor.RGB = bewertung.color
                '.Transparency = 0
            End With

            With .Fill
                .ForeColor.RGB = bewertung.color
                .ForeColor.TintAndShade = 0
                '.ForeColor.Brightness = 0.25
                .Transparency = 0.0

                'If roentgenBlick.isOn Then
                '    .Transparency = 0.8
                'Else
                '    If status = ProjektStatus(0) Then
                '        .Transparency = 0.35
                '    Else
                '        .Transparency = 0.0
                '    End If
                'End If

                .Solid()

            End With

            .TextFrame2.TextRange.Text = ""
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


        End With

    End Sub

    Public Sub defineShapeAppearance(ByRef myproject As clsProjekt, ByRef projectShape As Excel.Shape)
        Dim pcolor As Object
        Dim status As String
        Dim pMarge As Double
        Dim pname As String
        Dim diffToPrev As Boolean
        Dim ampel As Integer
        Dim showAmpel As Boolean = False
        Dim showResults As Boolean = True
        Dim myshape As Excel.Shape

        Try
            If projectShape.GroupItems.Count > 1 Then
                ' es handelt sich um die Darstellung inkl der Meilensteine
                myshape = projectShape.GroupItems(0)
            Else
                myshape = projectShape
            End If
        Catch ex As Exception
            myshape = projectShape
        End Try


        With myproject
            pcolor = .farbe
            status = .Status
            pMarge = .ProjectMarge
            pname = .name
            ampel = .ampelStatus
            diffToPrev = .diffToPrev
        End With

        With myshape

            If status = ProjektStatus(2) Or diffToPrev Then
                ' beauftragt, aber noch nicht wieder freigegeben ... 

                .Glow.Color.RGB = awinSettings.glowColor
                .Glow.Color.TintAndShade = 0
                .Glow.Color.Brightness = 0
                .Glow.Transparency = 0.4
                .Glow.Radius = 10

            Else
                .Glow.Color.RGB = RGB(255, 255, 255)
                .Glow.Transparency = 1.0
            End If

            With .Line
                .Visible = Microsoft.Office.Core.MsoTriState.msoTrue
                'If pMarge < 0 Then
                '    .ForeColor.RGB = RGB(255, 0, 0)
                '    .Weight = 2.0
                'Else
                '    .ForeColor.RGB = pcolor
                'End If
                .ForeColor.RGB = pcolor
                .Transparency = 0
            End With

            With .Fill
                '.Visible = msoTrue
                .ForeColor.RGB = pcolor
                .ForeColor.TintAndShade = 0
                .ForeColor.Brightness = -0.25

                If roentgenBlick.isOn Then
                    .Transparency = 0.8
                Else
                    .Transparency = 0.0
                End If

                .Solid()

            End With


            With .TextFrame2
                .VerticalAnchor = MsoVerticalAnchor.msoAnchorMiddle
                .HorizontalAnchor = MsoHorizontalAnchor.msoAnchorNone
            End With

            If roentgenBlick.isOn Then
                .TextFrame2.TextRange.Text = ""
            Else
                .TextFrame2.TextRange.Text = .Name
            End If

            If status = ProjektStatus(0) Then
                .Adjustments.Item(1) = 0.5
            Else
                .Adjustments.Item(1) = 0.25
            End If


        End With


    End Sub

    Public Sub defineShapeAppearance(ByRef myproject As clsProjekt, ByRef projectShape As Excel.Shape, ByVal phasenIndex As Integer)
        Dim projectColor As Object, phaseColor As Object = RGB(255, 255, 255)
        Dim status As String
        Dim pMarge As Double
        Dim pname As String
        Dim ampel As Integer
        Dim showAmpel As Boolean = False
        Dim showResults As Boolean = True
        Dim myshape As Excel.Shape
        Dim myphase As clsPhase

        Try
            myphase = myproject.getPhase(phasenIndex)

            Try
                phaseColor = myphase.Farbe
            Catch ex1 As Exception
                phaseColor = myproject.farbe
            End Try

        Catch ex As Exception
            Throw New ArgumentException("Phase " & phasenIndex.ToString & _
                                        " existiert nicht ...")
        End Try

        Try

            myshape = projectShape.GroupItems(phasenIndex - 1)

        Catch ex As Exception
            myshape = projectShape
        End Try


        With myproject
            projectColor = .farbe
            status = .Status
            pMarge = .ProjectMarge
            pname = .name
            ampel = .ampelStatus
        End With

        With myshape

            If status = ProjektStatus(2) Then

                If phasenIndex = 1 Then
                    ' beauftragt, aber noch nicht wieder freigegeben ... 

                    .Glow.Color.RGB = awinSettings.glowColor
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

            With .Line
                '.Visible = Microsoft.Office.Core.MsoTriState.msoTrue
                .Visible = Microsoft.Office.Core.MsoTriState.msoFalse
                'If pMarge < 0 Then
                '    .ForeColor.RGB = RGB(255, 0, 0)
                '    .Weight = 2.0
                'Else
                '    .ForeColor.RGB = pcolor
                'End If
                'If phasenIndex = 1 Then
                '    .ForeColor.RGB = projectColor
                '    .Transparency = 0
                'Else
                '    .ForeColor.RGB = phaseColor
                '    .Transparency = 0
                'End If

            End With

            With .Fill

                If phasenIndex = 1 Then
                    .ForeColor.RGB = projectColor
                Else
                    .ForeColor.RGB = phaseColor
                End If

                .ForeColor.TintAndShade = 0
                .ForeColor.Brightness = -0.25

                If roentgenBlick.isOn Then
                    .Transparency = 0.8
                Else
                    .Transparency = 0.0
                End If

                .Solid()

            End With




            If roentgenBlick.isOn Then
                .TextFrame2.TextRange.Text = ""
            Else
                If phasenIndex = 2 Then
                    With .TextFrame2
                        .VerticalAnchor = MsoVerticalAnchor.msoAnchorMiddle
                        .HorizontalAnchor = MsoHorizontalAnchor.msoAnchorNone
                        .WordWrap = MsoTriState.msoFalse
                    End With
                    .TextFrame2.TextRange.Text = pname
                Else
                    .TextFrame2.TextRange.Text = ""
                End If
            End If

            If status = ProjektStatus(0) Then
                .Adjustments.Item(1) = 0.5
            Else
                .Adjustments.Item(1) = 0.25
            End If


        End With


    End Sub

    ''' <summary>
    ''' passt die Shape Darstellung dem veränderten Projekt pname an  
    ''' </summary>
    ''' <remarks></remarks>
    Public Sub updateShapeinPlantafel(ByVal pname As String)
        Dim eeWasTrue As Boolean = False
        Dim suWasTrue As Boolean = False
        Dim zeile As Integer, spalte As Integer
        Dim laenge As Integer
        Dim status As String
        Dim top As Double, left As Double, width As Double, height As Double
        Dim magicBoardShapes As Excel.Shapes = appInstance.Worksheets(arrWsNames(3)).shapes
        Dim shpelement As Excel.Shape



        Dim hproj As New clsProjekt


        ' bestimmen der X- bzw. Y Position in der Plantafel 
        Try
            hproj = ShowProjekte.getProject(pname)
            With hproj
                zeile = .tfZeile
                spalte = .tfspalte
                laenge = .Dauer
                status = .Status
            End With
        Catch ex As Exception
            Call MsgBox("Fehler in clearProjektinPlantafel (Auslesen XPos, YPos, Dauer) von " & pname)
            Exit Sub
        End Try

        '
        ' hier wird in Plan Tafel das entsprechende Shape von der Erscheinung angepasst, ggf auch auf eine neue Zeile gesetzt ... 
        '
        Dim formerEE As Boolean = appInstance.EnableEvents
        Dim formerSU As Boolean = appInstance.ScreenUpdating
        appInstance.EnableEvents = False
        appInstance.ScreenUpdating = False

        With appInstance.Worksheets(arrWsNames(3))
            Try
                shpelement = magicBoardShapes.Item(pname)
                Dim myCollection As New Collection
                myCollection.Add(pname)
                zeile = findeMagicBoardPosition(myCollection, pname, zeile, spalte, laenge)

                ' jetzt ist eine passende Position gefunden ... die zugehörigen Shape Koordinaten werden berechnet 
                With hproj
                    .tfZeile = zeile
                    .CalculateShapeCoord(top, left, width, height)
                End With

                With shpelement
                    .Top = top
                    .Left = left
                    .Width = width
                    .Height = height
                End With

            Catch ex As Exception
                appInstance.EnableEvents = formerEE
                appInstance.ScreenUpdating = formerSU
                Throw New ArgumentException("updateProjektinPlantafel: kein Shape für Projekt " & pname & " gefunden")
            End Try


        End With

        appInstance.EnableEvents = formerEE
        appInstance.ScreenUpdating = formerSU


    End Sub


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
        Dim tmpshapes As Excel.Shapes = appInstance.Worksheets(arrWsNames(3)).shapes
        Dim shpelement As Excel.Shape

        Dim formerEE As Boolean = appInstance.EnableEvents
        appInstance.EnableEvents = False


        Dim formerSU As Boolean = appInstance.ScreenUpdating
        appInstance.ScreenUpdating = False

        ' Lösche das Shape Element
        Try
            shpelement = tmpshapes.Item(pname)
            With shpelement
                .Delete()
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
        appInstance.ScreenUpdating = formerSU


    End Sub
    ''' <summary>
    ''' zeichnet den Pfeil, der anzeigt, um wieviel ein Projekt bei Optimierung verschoben werden würde
    ''' </summary>
    ''' <param name="pname"></param>
    ''' <remarks></remarks>
    Public Sub ZeichneMoveLineOfProjekt(ByRef pname As String)

        Dim start As Integer
        Dim laenge As Integer
        Dim pcolor As Object, schriftfarbe As Object, fillColor As Object, borderColor As Object
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
            laenge = .Dauer
            start = .Start + .StartOffset
            moveLength = .StartOffset
            pcolor = .farbe
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
                    shp = .Shapes.AddShape(Type:=Microsoft.Office.Core.MsoAutoShapeType.msoShapeLeftArrow, _
                                left:=left, top:=top, width:=width, height:=height)

                Else
                    shp = .Shapes.AddShape(Type:=Microsoft.Office.Core.MsoAutoShapeType.msoShapeRightArrow, _
                                left:=left, top:=top, width:=width, height:=height)
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
        Call awinZeichnePlanTafel()
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
            allShapes = appInstance.ActiveSheet.shapes
        Catch ex As Exception
            allShapes = Nothing
        End Try

        If Not allShapes Is Nothing Then


            If calledFromPf Then
                ' der Name muss jetzt um das (xy.z%) bereinigt werden
                tmparray = pname.Split(New Char() {"("}, 10)
                Dim i As Integer
                For i = 0 To UBound(tmparray) - 1
                    realname = realname & tmparray(i)
                Next
                pname = realname.Trim
            End If

            If ShowProjekte.Liste.ContainsKey(pname) Then
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
            With appInstance.Worksheets(arrWsNames(3))
                For Each chtobj In .chartobjects
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
                            chartPT = .points(ptNr)
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



    ''' <summary>
    ''' speichert alle Projekte, die aktuell in Show- bzw NoShow sind
    ''' </summary>
    ''' <remarks></remarks>
    Public Sub awinExportAllProjects()
        '
        Dim dateinameQ As String, dateinameZ As String


        appInstance.ScreenUpdating = False
        Try
            ' hier muss jetzt das File Projekt Detail aufgemacht werden ...
            appInstance.Workbooks.Open(awinPath & projektAustausch)


            For Each kvp As KeyValuePair(Of String, clsProjekt) In AlleProjekte

                Try

                    Call awinExportProject(kvp.Value)

                Catch ex As Exception

                End Try


            Next kvp

            For Each kvp As KeyValuePair(Of String, clsProjekt) In DeletedProjekte.Liste
                dateinameQ = awinPath & projektFilesOrdner & "\" & kvp.Key & ".xlsm"
                dateinameZ = awinPath & deletedFilesOrdner & "\" & kvp.Key & ".xlsm"
                Try
                    My.Computer.FileSystem.MoveFile(dateinameQ, dateinameZ, True)
                Catch ex As Exception

                End Try


            Next kvp

        Catch ex As Exception
            Call MsgBox(ex.Message)
            Throw New ArgumentException("Abbruch - es konnten nicht ale Projekte gesichert werden ...")
            Exit Sub
        End Try

        appInstance.ActiveWorkbook.Close()
        appInstance.ScreenUpdating = True

    End Sub


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
        fileName = hproj.name & ".xlsx"

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


                ' Blattschutz setzen
                .Protect(Password:="x", UserInterfaceOnly:=True, DrawingObjects:=True, Contents:=True, Scenarios:=True)

            End With
        Catch ex As Exception
            ' Blattschutz setzen
            appInstance.Protect(Password:="x", UserInterfaceOnly:=True, DrawingObjects:=True, Contents:=True, Scenarios:=True)

            appInstance.EnableEvents = formerEE
            Throw New ArgumentException("Fehler in awinExportProject, Schreiben Stammdaten")
        End Try

        ' --------------------------------------------------
        ' jetzt werden die Ressourcen Bedarfe weggeschrieben 

        ' --------------------------------------------------

        Try
            With appInstance.ActiveWorkbook.Worksheets("Ressourcen")

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
                    .range("Zeitleiste").Cells(columnOffset).value = "= StartDatum"

                    .range("Zeitleiste").Cells(columnOffset + 1).value = "= EDATUM(D" & rowOffset & ",1"
                    .range("Zeitleiste").Cells(columnOffset + 2).value = "= EDATUM(E" & rowOffset & ",1"

                    ' die ersten beiden Felder der Zeitleiste formatieren
                    rng = .Range(.Cells(rowOffset, columnOffset + 1), .Cells(rowOffset, columnOffset + 2))
                    rng.NumberFormat = "mmm-yy"
                    ' Die restliche Zeitleiste  formatieren
                    'rng = .range(.cells(startZeile, spalte), .cells(endZeile, spalte))
                    destinationRange = .range(.Cells(rowOffset, columnOffset + 1), .Cells(rowOffset, columnOffset + 200))
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
                Dim itemName As String
                Dim dimension As Integer

                ' evtl hier vorher prüfen, ob es eine Phase mit Name hproj.name oder hproj.vorlagenName gibt; wenn nein , 
                ' muss hier der Projektname mit farbiger Gesamtdauer stehen 

                rowOffset = 1
                columnOffset = 1

                If hproj.CountPhases = 0 Then
                    ' Projekt-Name eintragen, Dauer einfärben

                    .range("Phasen_des_Projekts").Cells(rowOffset, columnOffset).value = hproj.name
                    .range("Phasen_des_Projekts").Cells(rowOffset, columnOffset).Interior.Color = hproj.farbe
                    rng = .Range("Zeitmatrix")(.Cells(rowOffset, columnOffset), .Cells(rowOffset, columnOffset + hproj.Dauer - 1))
                    rng.Interior.Color = hproj.farbe
                    rowOffset = rowOffset + 1
                End If

                For p = 1 To hproj.CountPhases
                    cphase = hproj.getPhase(p)
                    ' Phasen-Name eintragen, Dauer einfärben
                    itemName = cphase.name

                    Try
                        phasenFarbe = cphase.Farbe
                    Catch ex As Exception
                        phasenFarbe = hproj.farbe
                    End Try

                    If itemName = hproj.name Or itemName = hproj.VorlagenName Then
                        ' Projekt-Name eintragen, Dauer einfärben
                        .range("Phasen_des_Projekts").Cells(rowOffset, columnOffset).value = hproj.name
                        .range("Phasen_des_Projekts").Cells(rowOffset, columnOffset).Interior.Color = hproj.farbe
                        For d = 1 To hproj.Dauer
                            .Range("Zeitmatrix").Cells(rowOffset, columnOffset + d - 1).Interior.Color = hproj.farbe
                        Next d

                        d = appInstance.WorksheetFunction.CountA(.range("Phasen_des_Projekts"))



                    Else
                        .range("Phasen_des_Projekts").Cells(rowOffset, columnOffset).value = itemName
                        For d = 1 To cphase.relEnde - cphase.relStart + 1
                            .Range("Zeitmatrix").Cells(rowOffset, cphase.relStart + d - 1).Interior.Color = phasenFarbe
                        Next d

                    End If

                    rowOffset = rowOffset + 1



                    anzahlItems = cphase.CountRoles


                    ' jetzt werden Rollen geschrieben 
                    For r = 1 To anzahlItems
                        itemName = cphase.getRole(r).name
                        dimension = cphase.getRole(r).getDimension
                        'ReDim values(cphase.relEnde - cphase.relStart)
                        ReDim values(dimension)
                        values = cphase.getRole(r).Xwerte
                        .range("RollenKosten_des_Projekts").Cells(rowOffset, columnOffset).value = itemName

                        For d = 1 To dimension + 1
                            .Range("Zeitmatrix").Cells(rowOffset, cphase.relStart + d - 1).Interior.Color = phasenFarbe
                            .Range("Zeitmatrix").Cells(rowOffset, cphase.relStart + d - 1).Value = values(d - 1)
                        Next d
                        rowOffset = rowOffset + 1
                    Next r


                    ' jetzt werden Kosten geschrieben 

                    anzahlItems = cphase.CountCosts

                    For k = 1 To anzahlItems
                        itemName = cphase.getCost(k).name
                        dimension = cphase.getCost(k).getDimension
                        ReDim values(dimension)
                        values = cphase.getCost(k).Xwerte
                        .range("RollenKosten_des_Projekts").Cells(rowOffset, columnOffset).value = itemName
                        For d = 1 To dimension + 1
                            .Range("Zeitmatrix").Cells(rowOffset, cphase.relStart + d - 1).Interior.Color = phasenFarbe
                            .Range("Zeitmatrix").Cells(rowOffset, cphase.relStart + d - 1).Value = values(d - 1)
                        Next d
                        rowOffset = rowOffset + 1
                    Next
                    rowOffset = rowOffset + 1
                Next p

                ' Blattschutz setzen
                .Protect(Password:="x", UserInterfaceOnly:=True, DrawingObjects:=True, Contents:=True, Scenarios:=True)

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
                rng = .Range(.Cells(zeile, spalte), .Cells(zeile + 2000, spalte + 120))
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
                        rng = .cells(startZeile, spalte)
                    Else
                        rng = .range(.cells(startZeile, spalte), .cells(endZeile, spalte))
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
                        rng = .cells(startZeile, spalte)
                    Else
                        rng = .range(.cells(startZeile, spalte), .cells(endZeile, spalte))
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
                        rng = .cells(startZeile, spalte)
                    Else
                        rng = .range(.cells(startZeile, spalte), .cells(endZeile, spalte))
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
                    rng = .range(.cells(startKosten, spalte), .cells(endZeile, spalte))
                    appInstance.ActiveWorkbook.Names.Add(Name:="Kosten", RefersTo:=rng)

                End If

                If endZeile >= startZeile Then
                    rng = .range(.cells(startZeile, spalte), .cells(endZeile, spalte))
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
                    rng = .range(.cells(startZeile, spalte), .cells(endZeile, spalte))
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
            Dim cResult As New clsResult(parent:=cphase)
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

                For r = 1 To cphase.CountResults
                    cResult = cphase.getResult(r)

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

            ' Blattschutz setzen
            .Protect(Password:="x", UserInterfaceOnly:=True, DrawingObjects:=True, Contents:=True, Scenarios:=True)

        End With


        ' ----------------------------------------------
        ' jetzt werden die Attribute weggeschrieben ....

        Try
            With appInstance.ActiveWorkbook.Worksheets("Attribute")

                .Unprotect(Password:="x")       ' Blattschutz aufheben


                ' Projekt-Typ

                .range("Projekt_Typ").value = hproj.VorlagenName
                rng = .range("Projekt_Typ")
                With rng
                    .HorizontalAlignment = Excel.XlHAlign.xlHAlignLeft
                    .IndentLevel = 1
                    .WrapText = False
                End With

                ' Status

                .range("Status").value = hproj.Status
                rng = .range("Status")
                With rng
                    .HorizontalAlignment = Excel.XlHAlign.xlHAlignRight
                    .IndentLevel = 1
                    .WrapText = False
                End With

                ' Business_Unit

                .range("Business_Unit").value = hproj.businessUnit
                rng = .range("Business_Unit")
                With rng
                    .HorizontalAlignment = Excel.XlHAlign.xlHAlignLeft
                    .IndentLevel = 1
                    .WrapText = False
                End With

                ' Strategischer Fit

                .range("Strategischer_Fit").value = hproj.StrategicFit
                rng = .range("Strategischer_Fit")
                With rng
                    .HorizontalAlignment = Excel.XlHAlign.xlHAlignRight
                    .IndentLevel = 1
                    .WrapText = False
                End With

                ' Risiko

                .range("Risiko").value = hproj.Risiko
                rng = .range("Risiko")
                With rng
                    .HorizontalAlignment = Excel.XlHAlign.xlHAlignRight
                    .IndentLevel = 1
                    .WrapText = False
                End With


                ' Blattschutz setzen
                .Protect(Password:="x", UserInterfaceOnly:=True, DrawingObjects:=True, Contents:=True, Scenarios:=True)

            End With
        Catch ex As Exception
            ' Blattschutz setzen
            appInstance.ActiveWorkbook.Worksheets("Attribute").Protect(Password:="x", UserInterfaceOnly:=True, DrawingObjects:=True, Contents:=True, Scenarios:=True)
            appInstance.EnableEvents = formerEE
            Throw New ArgumentException("Fehler in awinExportProject, Schreiben Attribute")
        End Try



        Try
            My.Computer.FileSystem.DeleteFile(awinPath & projektFilesOrdner & "\" & fileName)
        Catch ex As Exception

        End Try

        Try
            appInstance.ActiveWorkbook.SaveAs(awinPath & projektFilesOrdner & "\" & fileName, _
                                          ConflictResolution:=XlSaveConflictResolution.xlLocalSessionChanges
                                          )
        Catch ex As Exception

        End Try


        appInstance.EnableEvents = formerEE


    End Sub

    ' Vorbedingung: das Active-workbook ist bereits das ProjektDetail File 
    Public Sub awinStoreProjForEditRess(hproj As clsProjekt)
        Dim rng As Excel.Range
        Dim zeile As Integer, spalte As Integer
        Dim delimiter As String = "."


        Dim pstart As Integer = hproj.Start

        With appInstance.Worksheets(arrWsNames(5))
            ' hier wird die erste Zeile beschrieben 

            For i = 1 To maxProjektdauer
                '.cells(1, i + 2).value = StartofCalendar.AddMonths(pstart + i - 2).ToString("MMM yy")
                .cells(1, i + 2).value = StartofCalendar.AddMonths(pstart + i - 2)
            Next i

            rng = .range(.cells(1, 3), .cells(1, maxProjektdauer + 2))

            Try
                rng.Columns.AutoFit()
            Catch ex As Exception

            End Try


            zeile = 2
            spalte = 1
            rng = .Range(.Cells(zeile, spalte), .Cells(zeile + 2000, spalte + 120))
            rng.Clear()

            Dim k As Integer
            k = 0
            ' wenn es noch keine Phasen gibt: Projekt-Name eintragen, Dauer einfärben
            If hproj.CountPhases = 0 Then
                ' Projekt-Name eintragen, Dauer einfärben
                .cells(zeile, spalte).value = hproj.name
                rng = .Range(.Cells(zeile, spalte + 2), .Cells(zeile, spalte + 1 + hproj.Dauer))
                rng.Interior.Color = hproj.farbe
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


                If itemName = hproj.name Or itemName = hproj.VorlagenName Then
                    ' Projekt-Name eintragen, Dauer einfärben
                    .cells(zeile, spalte).value = hproj.name
                    rng = .Range(.Cells(zeile, spalte + 2), .Cells(zeile, spalte + 1 + hproj.Dauer))
                    rng.Interior.Color = hproj.farbe
                    zeile = zeile + 1
                Else
                    ' Phasen Name eintragen, Dauer der Phase einfärben
                    .cells(zeile, spalte).value = itemName
                    rng = .Range(.Cells(zeile, spalte + 1 + cphase.relStart), .Cells(zeile, spalte + 1 + cphase.relEnde))
                    rng.Interior.Color = phasenFarbe
                    zeile = zeile + 1
                End If



                anzahlItems = cphase.CountRoles


                For r = 1 To anzahlItems
                    itemName = cphase.getRole(r).name
                    dimension = cphase.getRole(r).getDimension
                    ReDim values(dimension)
                    values = cphase.getRole(r).Xwerte
                    .cells(zeile, spalte + 1).value = itemName
                    rng = .Range(.Cells(zeile, spalte + 1 + cphase.relStart), .Cells(zeile, spalte + 1 + cphase.relStart + dimension))
                    rng.Value = values
                    zeile = zeile + 1
                Next r


                ' jetzt werden Kosten geschrieben 

                anzahlItems = cphase.CountCosts

                For k = 1 To anzahlItems
                    itemName = cphase.getCost(k).name
                    dimension = cphase.getCost(k).getDimension
                    ReDim values(dimension)
                    values = cphase.getCost(k).Xwerte
                    .cells(zeile, spalte + 1).value = itemName
                    rng = .Range(.Cells(zeile, spalte + 1 + cphase.relStart), .Cells(zeile, spalte + 1 + cphase.relStart + dimension))
                    rng.Value = values
                    zeile = zeile + 1
                Next k
                zeile = zeile + 1
            Next p

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
            lastRow = System.Math.Max(.cells(2000, 1).End(XlDirection.xlUp).row, .cells(2000, 2).End(XlDirection.xlUp).row) + 1
            rng = .range(.cells(2, 1), .cells(lastRow, 1))
            'If .cells(zeile, 1).value <> hproj.name Then
            '    hproj.name = .cells(zeile, 1).value
            'End If

            'zeile = 3
            For Each zelle In rng
                Select Case chkPhase
                    Case True
                        ' hier wird die Phasen Information ausgelesen

                        If Len(CType(zelle.Value, String)) > 1 Then
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
                                    .name = phaseName
                                    ' Änderung 28.11.13: jetzt wird die Phasen Länge exakt bestimmt , über startoffset in Tagen und dauerinDays als Länge
                                    Dim startOffset As Integer
                                    Dim dauerIndays As Integer
                                    startOffset = DateDiff(DateInterval.Day, hproj.startDate, hproj.startDate.AddMonths(anfang - 1))
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
                                    r = RoleDefinitions.getRoledef(hname).UID

                                    ReDim Xwerte(ende - anfang)


                                    'valueRange = .Range(zelle.Offset(0, anfang + 1), zelle.Offset(0, ende + 1))
                                    'Xwerte = CType(valueRange.Value, Double())

                                    For m = anfang To ende
                                        Xwerte(m - anfang) = zelle.Offset(0, m + 1).Value
                                    Next m

                                    crole = New clsRolle(ende - anfang)
                                    With crole
                                        .RollenTyp = r
                                        .Xwerte = Xwerte
                                    End With

                                    With cphase
                                        .AddRole(crole)
                                    End With
                                Catch ex As Exception
                                    '
                                    ' handelt es sich um die Kostenart Definition?
                                    ' 


                                End Try

                            ElseIf CostDefinitions.Contains(hname) Then

                                Try

                                    k = CostDefinitions.getCostdef(hname).UID

                                    ReDim Xwerte(ende - anfang)

                                    'valueRange = .Range(zelle.Offset(0, anfang + 1), zelle.Offset(0, ende + 1))
                                    'Xwerte = valueRange.Value

                                    For m = anfang To ende
                                        Xwerte(m - anfang) = zelle.Offset(0, m + 1).Value
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



    Public Sub awinReadProjectTemplate(ByVal pname As String, ByVal intern As Boolean)


        Dim lastRow As Integer
        Dim rng As Excel.Range
        Dim zelle As Excel.Range
        Dim zeile As Integer, spalte As Integer
        Dim hproj As New clsProjektvorlage

        zeile = 1
        spalte = 1


        Try
            With appInstance.ActiveWorkbook.Worksheets("General Information")

                hproj.VorlagenName = CType(.cells(zeile, spalte + 1).value, String).Trim
                hproj.Schrift = .cells(zeile, spalte + 1).font.size
                hproj.Schriftfarbe = .cells(zeile, spalte + 1).font.color
                hproj.farbe = .cells(zeile, spalte + 1).interior.color

                ' earliest
                hproj.earliestStart = -6
                ' latest
                hproj.latestStart = 6


            End With
        Catch ex As Exception
            Throw New ArgumentException("Fehler beim auslesen General Information")
        End Try


        Try
            With appInstance.ActiveWorkbook.Worksheets("Project Needs")

                Dim chkPhase As Boolean = True
                Dim Xwerte As Double()
                Dim crole As clsRolle
                Dim cphase As New clsPhase(hproj, True)
                Dim ccost As clsKostenart
                Dim phaseName As String
                Dim anfang As Integer, ende As Integer
                Dim farbeAktuell As Object
                Dim r As Integer, k As Integer
                'Dim valueRange As Excel.Range

                zeile = 2

                lastRow = System.Math.Max(.cells(2000, 1).End(XlDirection.xlUp).row, .cells(2000, 2).End(XlDirection.xlUp).row) + 1
                rng = .range(.cells(2, 1), .cells(lastRow, 1))

                For Each zelle In rng
                    Select Case chkPhase
                        Case True
                            ' hier wird die Phasen Information ausgelesen

                            If Len(CType(zelle.Value, String)) > 1 Then
                                phaseName = CType(zelle.Value, String).Trim

                                If Len(phaseName) > 0 Then

                                    cphase = New clsPhase(hproj, True)

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
                                        .name = phaseName
                                        ' Änderung 28.11.13: jetzt wird die Phasen Länge exakt bestimmt , über startoffset in Tagen und dauerinDays als Länge
                                        Dim startOffset As Integer = DateDiff(DateInterval.Day, StartofCalendar, StartofCalendar.AddMonths(anfang - 1))
                                        'Dim dauerIndays As Integer = DateDiff(DateInterval.Day, StartofCalendar.AddMonths(anfang - 1), _
                                        '                                                        StartofCalendar.AddMonths(ende).AddDays(-1)) + 1
                                        Dim dauerIndays As Integer = calcDauerIndays(StartofCalendar.AddDays(startOffset), ende - anfang + 1, True)
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
                                        r = RoleDefinitions.getRoledef(hname).UID

                                        ReDim Xwerte(ende - anfang)


                                        For m = anfang To ende
                                            Xwerte(m - anfang) = zelle.Offset(0, m + 1).Value
                                        Next m

                                        crole = New clsRolle(ende - anfang)
                                        With crole
                                            .RollenTyp = r
                                            .Xwerte = Xwerte
                                        End With

                                        With cphase
                                            .AddRole(crole)
                                        End With
                                    Catch ex As Exception
                                        '
                                        ' handelt es sich um die Kostenart Definition?
                                        ' 


                                    End Try

                                ElseIf CostDefinitions.Contains(hname) Then

                                    Try

                                        k = CostDefinitions.getCostdef(hname).UID

                                        ReDim Xwerte(ende - anfang)

                                        For m = anfang To ende
                                            Xwerte(m - anfang) = zelle.Offset(0, m + 1).Value
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
                    zeile = zeile + 1
                Next zelle



            End With
        Catch ex As Exception
            Throw New ArgumentException("Fehler in awinImportProject, Lesen Project Needs")
        End Try


        ' hier werden die mit den Phasen verbundenen Results ausgelesen ...

        Try
            With appInstance.ActiveWorkbook.Worksheets("Settings")
                rng = .Range("Phasen")
                Dim rngZeile As Excel.Range
                Dim lastColumn As Integer
                Dim resultName As String = ""
                Dim phaseName As String
                Dim tmpPhase As New clsPhase(hproj, True)
                Dim tmpStr() As String
                Dim defaultOffset As Integer


                Dim anzTage As Integer

                For Each zelle In rng

                    Try
                        phaseName = zelle.Value.trim

                        tmpPhase = hproj.getPhase(phaseName)
                        defaultOffset = tmpPhase.dauerInDays
                    Catch ex As Exception

                    End Try

                    If Not tmpPhase Is Nothing Then

                        rngZeile = rng.Rows(zelle.Row)
                        lastColumn = .cells(zelle.Row, 2000).End(XlDirection.xlToLeft).column

                        Dim specified As Boolean
                        For i = 4 To lastColumn

                            specified = False
                            Try
                                resultName = .cells(zelle.Row, i).value.ToString.Trim

                                tmpStr = resultName.Split(New Char() {"(", ")"}, 10)

                                If tmpStr.Length > 1 Then

                                    Try
                                        If awinSettings.offsetEinheit = "d" Then
                                            anzTage = CType(tmpStr(1), Integer)
                                        Else
                                            anzTage = CType(tmpStr(1), Integer) * 7
                                        End If

                                        resultName = tmpStr(0).Trim
                                        specified = True
                                    Catch ex1 As Exception
                                        resultName = .cells(zelle.Row, i).value.ToString.Trim
                                        anzTage = defaultOffset
                                    End Try

                                End If


                                Dim tmpResult As New clsResult(parent:=tmpPhase)

                                If resultName.Length > 0 Then
                                    With tmpResult
                                        .name = resultName
                                        If specified Then
                                            .offset = anzTage
                                        Else
                                            .offset = defaultOffset
                                        End If
                                    End With

                                    tmpPhase.AddResult(tmpResult)

                                End If
                            Catch ex As Exception

                            End Try
                        Next

                    End If

                Next

            End With
        Catch ex As Exception

        End Try



        Projektvorlagen.Add(hproj)


    End Sub



    Public Function textZeitraum(start As Integer, ende As Integer) As String
        Dim htxt As String = " "
        Dim von As Date, bis As Date

        If start <= 0 Then
            start = 1
        End If
        If ende > 120 Then
            ende = 120
        End If

        Try
            With appInstance.Worksheets(arrWsNames(3))
                von = .cells(1, start).value
                bis = .cells(1, ende).value
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

    Public Function getColumnOfDate(ByVal datum As Date) As Integer
        Dim spalte As Integer = 1

        Select Case awinSettings.zeitEinheit
            Case "PM"
                spalte = DateDiff(DateInterval.Month, StartofCalendar, datum) + 1
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
                    .Start + .Dauer - 1 >= getColumnOfDate(Date.Now) And _
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
    ''' errechnet die Kennung, die dem Chart als Namen mitgegeen wird; darf nicht größer als 31 in der Länge sein; 
    ''' das erlaubt Excel.Chart.name nicht
    ''' in typ wird mitgegeben , ob es sich um ein Portfolio Chart oder um ein Projekt Chart handelt 
    ''' 
    ''' </summary>
    ''' <param name="typ">ist Portfolio Chart (pf) oder Projekt-Chart (pr)</param>
    ''' <param name="index">gibt den Enumeration Wert an</param>
    ''' <param name="mycollection">enthält die Namen der Phasen/Rollen/Kostenarten</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function getKennung(ByVal typ As String, ByVal index As Integer, ByVal mycollection As Collection) As String
        Dim IDkennung As String
        Dim cName As String

        IDkennung = typ & "#" & index.ToString


        Try
            Select Case index
                Case PTpfdk.Phasen

                    If mycollection.Count = PhaseDefinitions.Count Then
                        IDkennung = IDkennung & "#Alle"

                    Else

                        For i = 1 To mycollection.Count
                            cName = mycollection.Item(i)
                            IDkennung = IDkennung & "#" & PhaseDefinitions.getPhaseDef(cName).UID.ToString
                        Next

                    End If

                Case PTpfdk.Rollen

                    If mycollection.Count = RoleDefinitions.Count Then
                        IDkennung = IDkennung & "#Alle"

                    Else

                        For i = 1 To mycollection.Count
                            cName = mycollection.Item(i)
                            IDkennung = IDkennung & "#" & RoleDefinitions.getRoledef(cName).UID.ToString
                        Next

                    End If

                Case PTpfdk.Kosten

                    If mycollection.Count = CostDefinitions.Count Then
                        IDkennung = IDkennung & "#Alle"

                    Else

                        For i = 1 To mycollection.Count
                            cName = mycollection.Item(i)
                            IDkennung = IDkennung & "#" & CostDefinitions.getCostdef(cName).UID.ToString
                        Next

                    End If

            End Select
        Catch ex As Exception
            IDkennung = IDkennung & "#?"
        End Try

        getKennung = IDkennung


    End Function

    Public Sub updateStatusInformation(ByVal resultShape As Excel.Shape)

        Dim tmpstr() As String
        Dim projectName As String
        Dim ok As Boolean = True
        Dim hproj As New clsProjekt
        Dim description As String

        With resultShape

            tmpstr = .Name.Split(New Char() {"#"}, 10)
            projectName = tmpstr(0)

            Try
                hproj = ShowProjekte.getProject(projectName)
            Catch ex As Exception
                hproj = Nothing
                ok = False
            End Try

            If ok Then

                Try

                    description = hproj.ampelErlaeuterung

                    With formStatus

                        .projectName.Text = hproj.name
                        .bewertungsText.Text = description

                        If .Visible Then
                        Else
                            .Visible = True
                            .Show()
                        End If

                    End With



                Catch ex As Exception

                    ok = False
                End Try

                If Not ok Then
                    'Call MsgBox("keine Information abrufbar ...")
                End If

            Else
                'Call MsgBox("keine Information abrufbar ...")
            End If


        End With

    End Sub

    Public Sub updateMilestoneInformation(ByVal resultShape As Excel.Shape)

        Dim tmpstr() As String
        Dim projectName As String
        Dim phaseName As String
        Dim cPhase As clsPhase
        Dim resultName As String
        Dim cResult As clsResult
        Dim bewertung As New clsBewertung
        Dim ok As Boolean = True
        Dim hproj As New clsProjekt


        With resultShape

            tmpstr = .Name.Split(New Char() {"#"}, 10)
            projectName = tmpstr(0)

            Try
                hproj = ShowProjekte.getProject(projectName)
            Catch ex As Exception
                hproj = Nothing
                ok = False
            End Try



            If ok Then

                Try
                    phaseName = tmpstr(1).Trim
                    cPhase = hproj.getPhase(phaseName)
                    'cResult = New clsResult(parent:=cPhase)
                    resultName = .Title
                    cResult = cPhase.getResult(resultName)

                    If IsNothing(cResult) Then
                    Else

                        With formMilestone
                            .bewertungsListe = cResult.bewertungsListe
                            .projectName.Text = hproj.name
                            .phaseName.Text = cPhase.name

                            .resultDate.Text = cResult.getDate.ToShortDateString
                            .resultName.Text = cResult.name


                            If .bewertungsListe.Count > 0 Then
                                Dim hb As clsBewertung = .bewertungsListe.ElementAt(0).Value

                                Dim farbe As System.Drawing.Color = System.Drawing.Color.FromArgb(hb.color)

                                .bewertungsText.Text = hb.description


                            Else

                                Dim farbe As System.Drawing.Color = System.Drawing.Color.FromArgb(awinSettings.AmpelNichtBewertet)

                                .bewertungsText.Text = "es existiert noch keine Bewertung ...."


                            End If

                            If .Visible Then
                            Else
                                .Visible = True
                                .Show()
                            End If

                        End With


                    End If




                Catch ex As Exception
                    phaseName = ""
                    resultName = ""
                    ok = False
                End Try

                If Not ok Then
                    'Call MsgBox("keine Information abrufbar ...")
                End If

            Else
                'Call MsgBox("keine Information abrufbar ...")
            End If


        End With



    End Sub

    Public Sub updatePhaseInformation(ByVal phaseShape As Excel.Shape)

        Dim tmpstr() As String
        Dim projectName As String
        Dim phaseName As String
        Dim cPhase As clsPhase

        Dim ok As Boolean = True
        Dim hproj As New clsProjekt

        Dim phaseStartdate As Date
        Dim phaseEnddate As Date
        Dim phaseDauerDays As Integer



        With phaseShape

            tmpstr = .Name.Split(New Char() {"#"}, 10)
            projectName = tmpstr(0)

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
                phaseName = tmpstr(1).Trim
                cPhase = hproj.getPhase(phaseName)
                'phaseStartdate = hproj.startDate.AddMonths(cPhase.relStart - 1)

                phaseStartdate = cPhase.getStartDate
                phaseEnddate = cPhase.getEndDate
                phaseDauerDays = cPhase.dauerInDays


                With formPhase

                    If specialListofPhases.Contains(phaseName) Then

                        .projectName.Text = projectName
                        .phaseName.Text = phaseName
                        .Height = 360
                        .erlaeuterung.Visible = True
                        .erlaeuterung.Text = " ... hier werden die versch. Register zu LeLe abrufbar sein ..."

                        .phaseStart.Text = phaseStartdate.ToShortDateString
                        .phaseStart.TextAlign = HorizontalAlignment.Left

                        .phaseEnde.Text = phaseEnddate.ToShortDateString
                        .phaseEnde.TextAlign = HorizontalAlignment.Right

                        .phaseDauer.Text = phaseDauerDays.ToString & " Tage"
                        .phaseDauer.TextAlign = HorizontalAlignment.Center


                        If .Visible Then
                        Else
                            .Visible = True
                            .Show()
                        End If


                    Else

                        .projectName.Text = projectName
                        .phaseName.Text = phaseName
                        .Height = 190
                        .erlaeuterung.Visible = False

                        .phaseStart.Text = phaseStartdate.ToShortDateString
                        .phaseStart.TextAlign = HorizontalAlignment.Left

                        .phaseEnde.Text = phaseEnddate.ToShortDateString
                        .phaseEnde.TextAlign = HorizontalAlignment.Right

                        .phaseDauer.Text = phaseDauerDays.ToString & " Tage"
                        .phaseDauer.TextAlign = HorizontalAlignment.Center

                        If .Visible Then
                        Else
                            .Visible = True
                            .Show()
                        End If


                    End If


                End With


            Catch ex As Exception
                phaseName = ""
                ok = False
            End Try


        Else
            'Call MsgBox("keine Information abrufbar ...")
        End If






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
        Dim phaseName As String


        For p = 1 To hproj.CountPhases

            Try
                If p = 1 Then
                    hphase = hproj.getPhase(1)
                    cphase = cproj.getPhase(1)

                    If hphase.startOffsetinDays = cphase.startOffsetinDays And _
                            hphase.dauerInDays = cphase.dauerInDays Then
                        Try
                            ' in diesem Fall müssen beide Phase(1) Namen, die ja evtl unterschiedlich sind, aufgenommen werden 
                            noColorCollection.Add(hphase.name, hphase.name)
                            noColorCollection.Add(cphase.name, cphase.name)
                        Catch ex As Exception

                        End Try
                    End If
                Else

                    hphase = hproj.getPhase(p)
                    phaseName = hphase.name

                    Try
                        cphase = cproj.getPhase(phaseName)

                        If hphase.startOffsetinDays = cphase.startOffsetinDays And _
                            hphase.dauerInDays = cphase.dauerInDays Then
                            Try
                                noColorCollection.Add(phaseName, phaseName)
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
    ''' Methode trägt alle Projekte aus ImportProjekte in AlleProjekte bzw. Showprojekte ein, sofern die Anzahl mit der myCollection übereinstimmt
    ''' die Projekte werden in der Reihenfolge auf das Board gezeichnet, wie sie in der myCollection aufgeführt sind
    ''' </summary>
    ''' <param name="myCollection"></param>
    ''' <param name="importDate"></param>
    ''' <remarks></remarks>
    Public Sub importProjekteEintragen(ByVal myCollection As Collection, ByVal importDate As Date)

        Dim hproj As New clsProjekt, cproj As New clsProjekt
        Dim pname As String, vglName As String

        Dim anzAktualisierungen As Integer, anzNeuProjekte As Integer
        Dim tafelZeile As Integer = 2

        Dim differentToPrevious As Boolean = False

        If myCollection.Count <> ImportProjekte.Count Then
            Throw New ArgumentException("keine Übereinstimmung in der Anzahl gültiger/ímportierter Projekte - Abbruch!")
        End If


        anzAktualisierungen = 0
        anzNeuProjekte = 0

        Dim ok As Boolean = True
        ' jetzt werden alle importierten Projekte bearbeitet 
        For Each pname In myCollection

            Try
                hproj = ImportProjekte.getProject(pname)
                pname = hproj.name

            Catch ex As Exception
                Call MsgBox("Projekt " & pname & " ist kein gültiges Projekt ... es wird ignoriert ...")
                ok = False
            End Try

            If ok Then

                ' jetzt muss überprüft werden, ob dieses Projekt bereits in AlleProjekte / Showprojekte existiert 
                ' wenn ja, muss es um die entsprechenden Werte dieses Projektes (Status, etc)  ergänzt werden
                ' wenn nein, wird es im Show-Modus ergänzt 

                vglName = hproj.name & "#" & hproj.variantName
                Try
                    cproj = AlleProjekte.Item(vglName)
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
                            If cproj.Erloes > 0 Then
                                ' dann soll der alte Wert beibehalten werden 
                                .Erloes = cproj.Erloes
                            End If

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

                            End If

                        End With
                    Catch ex As Exception
                        ok = False
                        Throw New ArgumentException("Fehler bei Übernahme der Attribute des alten Projektes" & vbLf & ex.Message)

                    End Try


                    Try
                        If ShowProjekte.Liste.ContainsKey(pname) Then


                            ' Shape wird auf der Plan-Tafel gelöscht - ausserdem wird der Verweis in hproj auf das Shape gelöscht 
                            Call clearProjektinPlantafel(hproj.name)

                            ShowProjekte.Remove(pname)


                        End If

                        AlleProjekte.Remove(vglName)

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

                            .earliestStart = 0
                            .earliestStartDate = .startDate

                            .Id = vglName & "#" & importDate.ToString
                            .latestStart = 0
                            .latestStartDate = .startDate
                            .leadPerson = " "
                            .shpUID = ""
                            .StartOffset = 0
                            .Status = ProjektStatus(0)
                            '.tfSpalte = 0
                            .tfZeile = tafelZeile
                            .timeStamp = importDate
                            .UID = cproj.UID

                        End With

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
                        ShowProjekte.Add(hproj)
                        AlleProjekte.Add(vglName, hproj)

                        ' ggf Bedarfe anzeigen 
                        If roentgenBlick.isOn Then
                            With roentgenBlick
                                Call awinShowNeedsofProject1(mycollection:=.myCollection, type:=.type, projektname:=pname)
                            End With

                        End If

                        ' zeichne das neue Shape in der Plan-Tafel 
                        Call ZeichneProjektinPlanTafel(pname, hproj.tfZeile, False)

                        ' jetzt müssen die ggf aktuell gezeigten Diagramme neu gezeichnet werden 
                        Call awinNeuZeichnenDiagramme(2)

                    Catch ex As Exception
                        Call MsgBox("Fehler bei Eintrag Showprojekte / Import " & hproj.name)
                    End Try

                End If


            End If

        Next

        If ImportProjekte.Count < 1 Then
            Call MsgBox(" es waren keine Projekte zu importieren ...")
        Else
            Call MsgBox("es wurden " & ImportProjekte.Count & " Projekte importiert!" & vbLf & _
                        anzNeuProjekte.ToString & " neue Projekte" & vbLf & _
                        anzAktualisierungen.ToString & " Projekt-Aktualisierungen")

        End If

        ImportProjekte.Clear()

    End Sub

    'Public Function zeichneShapeOfProject(ByRef hproj As clsProjekt, _
    '                                     ByVal zeile As Integer, _
    '                                     ByVal drawPhases As Boolean) As Excel.Shape
    '    Dim newShpElement As Excel.Shape
    '    Dim groupShpElement As Excel.Shape
    '    Dim top As Double, left As Double, width As Double, height As Double
    '    Dim pname As String
    '    Dim worksheetShapes As Excel.Shapes


    '    Try

    '        worksheetShapes = CType(appInstance.Worksheets(arrWsNames(3)), Excel.Worksheet).Shapes

    '    Catch ex As Exception
    '        Throw New Exception("in zeichneShapeOfProject : keine Shapes Zuordnung möglich ")
    '    End Try

    '    With hproj
    '        '.tfSpalte = start
    '        .tfZeile = zeile
    '        pname = .name

    '    End With


    '    If drawPhases Then

    '        Dim shapeGroupListe() As Object
    '        Dim anzGroupElemente As Integer = 0
    '        Dim shapesCollection As New Collection
    '        Dim phasenfarbe As Integer
    '        Dim phasenName As String
    '        Dim shapeName As String

    '        newShpElement = Nothing
    '        For i = 1 To hproj.CountPhases
    '            phasenName = hproj.getPhase(i).name
    '            phasenFarbe = hproj.getPhase(i).Farbe
    '            hproj.CalculateShapeCoord(i, top, left, width, height)



    '            Try
    '                If i = 1 Then
    '                    newShpElement = worksheetShapes.AddShape(Type:=Microsoft.Office.Core.MsoAutoShapeType.msoShapePentagon, _
    '                       Left:=left, Top:=top, Width:=width, Height:=height)
    '                Else
    '                    newShpElement = worksheetShapes.AddShape(Type:=Microsoft.Office.Core.MsoAutoShapeType.msoShapeChevron, _
    '                       Left:=left, Top:=top, Width:=width, Height:=height)
    '                End If
    '            Catch ex As Exception
    '                Throw New Exception("in zeichneShapeOfProject : keine Shape-Erstellung möglich ...  ")
    '            End Try


    '            shapeName = pname & "#" & phasenName & "#" & i.ToString
    '            With newShpElement
    '                .Name = shapeName
    '            End With

    '            Call defineShapeAppearance(hproj, newShpElement, i)


    '            Try
    '                shapesCollection.Add(shapeName, Key:=shapeName)
    '            Catch ex As Exception

    '            End Try


    '        Next

    '        ' hier werden die Shapes gruppiert
    '        anzGroupElemente = shapesCollection.Count

    '        If anzGroupElemente > 1 Then
    '            ' es macht nur Sinn zu gruppieren, wenn es mehr als 1 Element ist ....

    '            ReDim shapeGroupListe(anzGroupElemente - 1)
    '            For i = 1 To anzGroupElemente
    '                shapeGroupListe(i - 1) = shapesCollection.Item(i)
    '            Next

    '            Dim ShapeGroup As Excel.ShapeRange
    '            ShapeGroup = worksheetShapes.Range(shapeGroupListe)
    '            groupShpElement = ShapeGroup.Group()




    '            ShowProjekte.AddShape(pname, shpUID:=groupShpElement.ID.ToString)

    '        Else
    '            ' in diesem Fall besteht das Projekt nur aus einer einzigen Phase
    '            groupShpElement = newShpElement

    '        End If

    '        Try
    '            With groupShpElement
    '                .Name = pname
    '                hproj.shpUID = .ID.ToString
    '            End With
    '        Catch ex As Exception
    '            Throw New Exception("in zeichneShapeOfProject : dem shpae kann kein Name zugewiesen werden ....   ")
    '        End Try

    '        ' jetzt muss das neue Shape in der ShowProjekte.ShapeListe eingetragen werden ..
    '        ShowProjekte.AddShape(pname, shpUID:=groupShpElement.ID.ToString)



    '    Else

    '        hproj.CalculateShapeCoord(top, left, width, height)

    '        newShpElement = worksheetShapes.AddShape(Type:=Microsoft.Office.Core.MsoAutoShapeType.msoShapeRoundedRectangle, _
    '                Left:=left, Top:=top, Width:=width, Height:=height)


    '        With newShpElement
    '            .Name = pname
    '            hproj.shpUID = .ID.ToString
    '        End With

    '        Call defineShapeAppearance(hproj, newShpElement)

    '        'If showresults Then

    '        '    Call zeichneResultMilestonesInPlantafel(hproj)

    '        'End If

    '        ' jetzt muss das neue Shape in der ShowProjekte.ShapeListe eingetragen werden ..
    '        ShowProjekte.AddShape(pname, shpUID:=newShpElement.ID.ToString)
    '    End If




    '    zeichneShapeOfProject = newShpElement



    'End Function
End Module
