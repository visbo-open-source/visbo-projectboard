Imports Microsoft.Office.Interop
Imports Microsoft.Office.Interop.Excel
Imports System.Math

''' <summary>
''' Klasse , die eine Liste mit den Angaben zu eindeutiger Bezeichner und Koordinaten von allen Projekt-Shapes enthält 
''' wrid hauptsächlich benötigt, um festzustellen, wann sich ein Shape verschoben hat, gestaucht, gedehnt etc wurde 
''' Es sind auch Methoden enthalten, um verschobene Shapes mit ihrem Projekt Pendant zu synchronisieren
''' </summary>
''' <remarks></remarks>
''' 
Public Class clsProjektShapes

    Private AllShapes As SortedList(Of String, Double())
    

    ''' <summary>
    ''' gibt die Zeile zurück, ab der nach unten in der Projekttafel frei ist
    ''' maxzeile ist auch bereits frei 
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks>
    ''' shp.value(0) enthält top , shp.value(2) enthält height</remarks>
    Public ReadOnly Property getMaxZeile() As Integer
        Get
            Dim maxzeile As Integer = 1

            For Each shpElem As KeyValuePair(Of String, Double()) In AllShapes

                If shpElem.Value.Length > 3 Then
                    If CInt(1 + (shpElem.Value(0) + shpElem.Value(2) - topOfMagicBoard) / boxHeight) > maxzeile Then
                        maxzeile = CInt(1 + (shpElem.Value(0) + shpElem.Value(2) - topOfMagicBoard) / boxHeight)
                    End If
                End If

            Next

            getMaxZeile = maxzeile

        End Get
    End Property

    ''' <summary>
    ''' gibt true zurück, wenn der Name in der Shape Liste enthalten ist
    ''' false, sonst
    ''' </summary>
    ''' <param name="suchName"></param>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public ReadOnly Property contains(ByVal suchName As String) As Boolean
        Get

            Try
                contains = AllShapes.ContainsKey(suchName)
            Catch ex As Exception
                contains = False
            End Try


        End Get
    End Property


    ''' <summary>
    ''' gibt ein ShapeRange Objekt zurück, das alle Shapes enthält, die in Shape mit Namen pName enthalten sind 
    ''' </summary>
    ''' <param name="pName"></param>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public ReadOnly Property ungroupShapes(ByVal pName As String) As Excel.ShapeRange

        Get
            Dim projectShapes As Excel.Shapes
            Dim projectShape As Excel.ShapeRange

            ' hier sind alle Shapes drin
            projectShapes = CType(appInstance.Worksheets(arrWsNames(3)), Excel.Worksheet).Shapes

            ' hole das Projekt-Shape 
            projectShape = projectShapes.Range(pName)

            ungroupShapes = projectShape.Ungroup()
        End Get


    End Property


    ''' <summary>
    ''' gruppiert die im Array übergebenen Shapes zu einem Shape mit Namen pName
    ''' </summary>
    ''' <param name="listOFNames"></param>
    ''' <param name="pName"></param>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public ReadOnly Property groupShapes(ByVal listOFNames As Collection, ByVal pName As String) As Excel.Shape
        Get
            Dim worksheetShapes As Excel.Shapes
            Dim arrayOFNames() As String
            Dim i As Integer
            Dim anzElements As Integer = listOFNames.Count
            Dim shapegruppe As Excel.ShapeRange
            Dim hproj As clsProjekt

            Try
                hproj = ShowProjekte.getProject(pName)

                With CType(appInstance.Worksheets(arrWsNames(3)), Excel.Worksheet)
                    worksheetShapes = .Shapes
                End With



                If anzElements = 0 Then

                    groupShapes = Nothing

                ElseIf anzElements = 1 Then

                    groupShapes = worksheetShapes.Item(listOFNames.Item(1))

                Else

                    ReDim arrayOFNames(anzElements - 1)

                    For i = 1 To anzElements
                        arrayOFNames(i - 1) = CStr(listOFNames.Item(i))
                    Next

                    shapegruppe = worksheetShapes.Range(arrayOFNames)
                    groupShapes = shapegruppe.Group

                End If

                If anzElements > 0 Then
                    With groupShapes

                        .Name = pName
                        If awinSettings.drawphases Then
                            .AlternativeText = CInt(PTshty.projektE).ToString
                        Else
                            If anzElements = 1 Then
                                .AlternativeText = CInt(PTshty.projektN).ToString
                            Else
                                .AlternativeText = CInt(PTshty.projektC).ToString
                            End If

                        End If

                        hproj.shpUID = .ID.ToString

                    End With

                    ' jetzt muss das auch in der Liste Showprojekte eingetragen werden 
                    ShowProjekte.AddShape(pName, hproj.shpUID)

                    If anzElements > 1 Then
                        ' jetzt muss das Phase 1 Shape als shty.phaseN deklariert werden 
                        ' damit ändert sich auch der Name des Shapes 
                        Dim phase1Shape As Excel.Shape


                        Try
                            phase1Shape = groupShapes.GroupItems.Item(1)

                            With phase1Shape
                                If awinSettings.drawphases Then
                                    .AlternativeText = CInt(PTshty.phaseE).ToString
                                Else
                                    .AlternativeText = CInt(PTshty.phase1).ToString
                                End If

                                .Name = projectboardShapes.calcPhaseShapeName(hproj.name, hproj.getPhase(1).name)


                            End With



                        Catch ex As Exception

                        End Try

                    End If

                End If

            Catch ex As Exception
                Call MsgBox("Fehler in groupShapes: " & vbLf & ex.Message)
                groupShapes = Nothing
            End Try


            

        End Get
    End Property

    ''' <summary>
    ''' macht eine Re-Gruppierung der Shapesammlung
    ''' kann nur aufgerufen werden, wenn die Shapesammlung unverändert ist 
    ''' am 12.7 eine Public Property geworden
    ''' </summary>
    ''' <param name="shapeSammlung"></param>
    ''' <param name="pName"></param>
    ''' <remarks></remarks>
    Public Sub reGroupShape(ByRef shapeSammlung As Excel.ShapeRange, ByVal pName As String)
        Dim pShape As Excel.Shape
        Dim hproj As clsProjekt


        Try

            hproj = ShowProjekte.getProject(pName)
            pShape = shapeSammlung.Regroup

            With pShape
                .Name = pName
                .AlternativeText = CInt(PTshty.projektE).ToString
                hproj.shpUID = .ID.ToString

            End With

            ' jetzt muss das auch in der Liste Showprojekte eingetragen werden 
            ShowProjekte.AddShape(pName, hproj.shpUID)
        Catch ex As Exception

        End Try
        

    End Sub

    ''' <summary>
    ''' berechnet für das Projekt mit Namen pName den korrespondierenden Projekt-Shape-Namen und gibt ihn zurück  
    ''' </summary>
    ''' <param name="pName"></param>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public ReadOnly Property calcProjektShapeName(ByVal pName As String) As String
        Get
            calcProjektShapeName = pName
        End Get
    End Property


    ''' <summary>
    ''' berechnet für die Phase phaseName des Projekts pName den Shape-Namen und gibt ihn zurück
    ''' </summary>
    ''' <param name="pName"></param>
    ''' <param name="phaseName"></param>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public ReadOnly Property calcPhaseShapeName(ByVal pName As String, ByVal phaseName As String) As String
        Get
            calcPhaseShapeName = pName & "#" & phaseName
        End Get
    End Property

    ''' <summary>
    ''' berechnet für den Meilenstein mit laufender Nummer lfdNr in Phase phaseNAme in Projekt pName 
    ''' den Milestone-Shape Namen und gibt ihn zurück 
    ''' </summary>
    ''' <param name="pName">Projekt-Name</param>
    ''' <param name="phaseName">Phasen-Name</param>
    ''' <param name="lfdNr">lfd Nummer des Meilensteins in Phase phaseName</param>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public ReadOnly Property calcMilestoneShapeName(ByVal pName As String, ByVal phaseName As String, ByVal lfdNr As Integer) As String
        Get
            calcMilestoneShapeName = pName & "#" & phaseName & "#M" & lfdNr.ToString
        End Get
    End Property

    ''' <summary>
    ''' berechnet für das Projekt mit Namen pName den korrespondierenden Status-Shape Namen 
    ''' für den Monat, der in Kalenderspalte dateColumn liegt und gibt ihn zurück  
    ''' </summary>
    ''' <param name="dateColumn"></param>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public ReadOnly Property calcStatusShapeName(ByVal pName As String, ByVal dateColumn As Integer) As String
        Get
            'calcStatusShapeName = pName & "#Status#" & dateColumn.ToString
            ' der status soll nur einmal gezeichnet werden 
            calcStatusShapeName = pName & "#Status#"
        End Get
    End Property


    ''' <summary>
    ''' berechnet für die Projekte pname (project) und dpName (dependent project) den korrespondierenden Shape-Namen
    ''' </summary>
    ''' <param name="pName"></param>
    ''' <param name="dpName"></param>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public ReadOnly Property calcDependencyShapeName(ByVal pName As String, dpName As String) As String
        Get
            calcDependencyShapeName = pName.Trim & "#" & dpName.Trim
        End Get
    End Property

    ''' <summary>
    ''' fügt das Shape der Shapeslist hinzu; wenn es schon existiert, wird der alte Eintrag gelöscht und der neue 
    ''' wird eingetragen
    ''' </summary>
    ''' <param name="shpElement"></param>
    ''' <remarks></remarks>
    Public Sub add(ByVal shpElement As Excel.Shape)

        Dim key As String = shpElement.Name
        Dim shpCoord(3) As Double

        With shpElement
            shpCoord(0) = .Top
            shpCoord(1) = .Left
            shpCoord(2) = .Height
            shpCoord(3) = .Width
        End With

        If AllShapes.ContainsKey(key) Then
            ' existiert schon 
            AllShapes.Item(key) = shpCoord
        Else
            AllShapes.Add(key, shpCoord)
        End If

    End Sub

    ''' <summary>
    ''' löscht zu dem angegeben projektshape die Child-Shapes mit den in typCollection angegebenen Typen
    ''' wenn typCollection Null ist , dann sollen alle Elemente gelöscht werden 
    ''' Ausnahme: die Phase 1 darf nicht geöscht werden 
    ''' </summary>
    ''' <param name="projektshape"></param>
    ''' <param name="typCollection"></param>
    ''' <remarks></remarks>
    Public Sub removeChildsOfType(ByRef projektshape As Excel.Shape, ByVal typCollection As Collection)

        
        Dim pName As String

        Dim done As Boolean
        Dim shapeGruppe As ShapeRange
        Dim nameCollection As New Collection

        ' nur dann kann es aus mehreren bestehen ....
        If projektshape.AutoShapeType = Microsoft.Office.Core.MsoAutoShapeType.msoShapeMixed Then

            pName = extractName(projektshape.Name, PTshty.projektN)

            shapeGruppe = projektshape.Ungroup

            For Each elem As Excel.Shape In shapeGruppe

                If typCollection.Contains(elem.AlternativeText) Then
                    done = Me.AllShapes.Remove(elem.Name)
                    elem.Delete()
                Else
                    nameCollection.Add(elem.Name, elem.Name)
                End If

            Next

            If nameCollection.Count > 0 Then

                projektshape = Me.groupShapes(nameCollection, pName)

            Else

            End If

            Me.add(projektshape)

        End If


    End Sub
  
    ''' <summary>
    ''' gibt eine Collection der Shape-Namen zurück, die im gruppierten Shape projektshape enthalten sind
    ''' jedes Item hat folgende Struktur: 
    ''' PTshty.typ#Name des Shapes
    ''' </summary>
    ''' <param name="projektShape"></param>
    ''' <param name="typCollection">
    ''' leer: alle
    ''' sonst die Namen der Phasen bzw. Meilensteine, je nach typ 
    ''' </param>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public ReadOnly Property getAllChildswithType(ByVal projektShape As Excel.Shape, ByVal typCollection As Collection) As Collection
        Get
            Dim tmpCollection As New Collection
            Dim i As Integer
            Dim tmpShape As Excel.Shape
            Dim anzElements As Integer
            Dim hproj As clsProjekt = Nothing
            Dim ok As Boolean

            Try
                hproj = ShowProjekte.getProject(projektShape.Name)
                ok = True
            Catch ex As Exception
                ok = False
            End Try


            If ok Then

                If projektShape.AutoShapeType = Microsoft.Office.Core.MsoAutoShapeType.msoShapeMixed Then

                    Try
                        anzElements = projektShape.GroupItems.Count
                    Catch ex As Exception
                        anzElements = 0
                    End Try



                    For i = 1 To anzElements

                        tmpShape = projektShape.GroupItems.Item(i)


                        If typCollection.Count = 0 Then

                            If tmpShape.Name <> projektShape.Name And _
                                Not tmpCollection.Contains(tmpShape.AlternativeText & "#" & tmpShape.Name) Then

                                tmpCollection.Add(tmpShape.AlternativeText & "#" & tmpShape.Name, tmpShape.AlternativeText & "#" & tmpShape.Name)

                            End If

                        Else
                            Dim elementName As String

                            If tmpShape.Name <> projektShape.Name And typCollection.Contains(tmpShape.AlternativeText) Then

                                If tmpShape.AlternativeText = CInt(PTshty.phaseE).ToString Or _
                                    tmpShape.AlternativeText = CInt(PTshty.phaseN).ToString Then

                                    elementName = extractName(tmpShape.Name, PTshty.phaseN)

                                    If elementName <> projektShape.Name And Not tmpCollection.Contains(elementName) Then
                                        tmpCollection.Add(elementName, elementName)
                                    End If

                                ElseIf tmpShape.AlternativeText = CInt(PTshty.milestoneE).ToString Or _
                                        tmpShape.AlternativeText = CInt(PTshty.milestoneN).ToString Then

                                    Dim phaseName As String = extractName(tmpShape.Name, PTshty.phaseN)
                                    Dim msNr As Integer = CInt(extractName(tmpShape.Name, PTshty.milestoneN))


                                    Try
                                        elementName = hproj.getPhase(phaseName).getResult(msNr).name
                                        If Not tmpCollection.Contains(elementName) Then
                                            tmpCollection.Add(elementName, elementName)
                                        End If
                                    Catch ex As Exception

                                    End Try


                                End If

                            End If

                        End If


                    Next

                End If

            End If
            

            getAllChildswithType = tmpCollection

        End Get
    End Property

    ''' <summary>
    ''' gibt eine Collection der Shape-Namen zurück, die zu Projekt pName gehören 
    ''' jedes Item hat folgende Struktur: 
    ''' Name des Shapes
    ''' </summary>
    ''' <param name="pName"></param>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public ReadOnly Property getAllChilds(pName As String) As Collection
        Get

            Dim tmpCollection As New Collection

            For Each kvp As KeyValuePair(Of String, Double()) In Me.AllShapes
                If extractName(kvp.Key, PTshty.projektN) = pName And kvp.Key <> pName Then
                    tmpCollection.Add(kvp.Key)
                End If
            Next

            getAllChilds = tmpCollection

        End Get
    End Property


    ''' <summary>
    ''' entfernt das Shape aus der Shapesliste und löscht es von der Projekttafel
    ''' wenn es sich um ein projekt-Shape handelt, werden alle abhängigen Shapes wie Phasen, Meilensteine, Stati 
    ''' mitgelöscht
    ''' </summary>
    ''' <param name="shpElement">shpElement </param>
    ''' <remarks></remarks>
    Public Sub remove(ByRef shpElement As Excel.Shape)

        Dim shpName As String = shpElement.Name
        Dim done As Boolean
        'Dim todoListe1 As New Collection
        Dim todoListe2 As New Collection
        Dim shapetype As Integer = kindOfShape(shpElement)
        Dim tmpCollection As New Collection


        If isProjectType(shapetype) Then
            'Test
            'todoListe1 = Me.getAllChildswithType(shpElement, tmpCollection)
            todoListe2 = Me.getAllChilds(shpElement.Name)

            'Test 
            'If todoListe1.Count <> todoListe2.Count Then
            'Call MsgBox("Fehler in clsprojektShapes.remove)")
            'End If
        End If

        Try

            shpElement.Delete()
            done = Me.AllShapes.Remove(shpName)

        Catch ex As Exception

        End Try

        Try
            If todoListe2.Count > 0 Then
                Dim childName As String

                For Each childName In todoListe2

                    Try
                        done = Me.AllShapes.Remove(childName)
                    Catch ex2 As Exception

                    End Try

                Next

            End If
        Catch ex As Exception

        End Try



    End Sub

    ''' <summary>
    ''' liefert die Koordinaten des Shapes, wie sie noch in der Shapesliste stehen 
    ''' </summary>
    ''' <param name="shpName"></param>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public ReadOnly Property getCoord(ByVal shpName As String) As Double()
        Get
            getCoord = AllShapes.Item(shpName)
        End Get
    End Property

    ''' <summary>
    ''' prüft ob sich das Shape um mehr als die Toleranz-Werte verschoben hat 
    ''' wenn nein, werden die Werte wieder auf die alten zurückgesetzt 
    ''' </summary>
    ''' <param name="shpElement"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function hasAchanged(ByRef shpElement As Excel.Shape) As Boolean

        Dim key As String = shpElement.Name
        Dim tolX As Double = boxWidth * 0.06
        Dim tolY As Double = boxHeight * 0.7
        Dim shpCoord(3) As Double
        Dim oldCoord() As Double
        Dim isdifferent As Boolean = False
        Dim shapeType As Integer = kindOfShape(shpElement)

        Dim pName As String = ""
        Dim shapeSammlung As Excel.ShapeRange = Nothing

        ' Änderung 10.7.14 es gibt jetzt keinen alleinstehenden Meilenstein oder Phasen oder status Shape  
        'If shapeType = PTshty.phaseE Or shapeType = PTshty.milestoneE Then

        If Not isProjectType(shapeType) Then

            ' hier muss erst mal das übergeordnete Projektshape gesucht und ungruppiert werden  
            pName = extractName(shpElement.Name, PTshty.projektE)
            shapeSammlung = ungroupShapes(pName)

        End If

        With shpElement
            shpCoord(0) = .Top
            shpCoord(1) = .Left
            shpCoord(2) = .Height
            shpCoord(3) = .Width
        End With


        Try
            oldCoord = AllShapes.Item(key)

            ' Top überprüfen
            If Abs(oldCoord(0) - shpCoord(0)) > tolY Then
                isdifferent = True
            Else
                shpElement.Top = CSng(oldCoord(0))
            End If

            ' Left überprüfen
            If Abs(oldCoord(1) - shpCoord(1)) > tolX Then
                isdifferent = True
            Else
                shpElement.Left = CSng(oldCoord(1))
            End If

            ' Höhe darf niemals verändert werden ...
            shpElement.Height = CSng(oldCoord(2))

            ' Width überprüfen: width darf nur bei phaseE, phaseN, projektN, projektC, projektE
            If shapeType = PTshty.phaseE Or shapeType = PTshty.phaseN Or shapeType = PTshty.phase1 Or _
                isProjectType(shapeType) Then

                If Abs(oldCoord(3) - shpCoord(3)) > tolX * 0.2 Then
                    isdifferent = True
                Else
                    shpElement.Width = CSng(oldCoord(3))
                End If

            Else
                shpElement.Width = CSng(oldCoord(3))
            End If


            If Abs(shpElement.Rotation) > 15 Then
                isdifferent = True
            End If

        Catch ex As Exception
            isdifferent = False
        End Try


        'If shapeType = PTshty.phaseE Or shapeType = PTshty.milestoneE Then
        If Not isProjectType(shapeType) Then
            ' jetzt wird wieder regruppiert

            Call reGroupShape(shapeSammlung, pName)

        End If

        hasAchanged = isdifferent

    End Function



    Public Sub clear()
        AllShapes.Clear()
    End Sub

    ''' <summary>
    ''' prüft, ob das Shape noch an der alten Stelle steht
    ''' wenn nein: ist das Verschieben / Stauchen / Dehnen zulässig ? 
    ''' wenn ja:  das Projekt wird geändert, so dass es mit dem shape übereinstimmt   
    ''' </summary>
    ''' <param name="shpElement">ShapeELement</param>
    ''' <remarks></remarks>
    Public Sub sync(ByRef shpElement As Excel.Shape, ByVal selCollection As Collection)
        Dim curCoord(3) As Double, oldCoord() As Double
        Dim shapeType As Integer
        Dim moveAllowed As Boolean
        Dim phaseName As String, resultNr As Integer
        Dim hproj As clsProjekt, newProjekt As clsProjekt
        Dim tmpRange As Excel.ShapeRange
        Dim pShape As Excel.Shape

        Dim pName As String = ""
        Dim shapeSammlung As Excel.ShapeRange = Nothing
        ' wird benötigt, um festzustellen, ob noch eine Gruppierung am Ende gemacht werden muss 
        Dim notRegroupedAgain As Boolean = True


        Try
            shapeType = CInt(shpElement.AlternativeText)

            pName = extractName(shpElement.Name, PTshty.projektN)
            hproj = ShowProjekte.getProject(pName)

            If hproj.Status = ProjektStatus(0) _
                And Not shapeType = PTshty.status _
                And Not shapeType = PTshty.dependency _
                And Not shapeType = PTshty.phase1 Then
                moveAllowed = True
            Else
                moveAllowed = False
            End If

            oldCoord = AllShapes(key:=shpElement.Name)

        Catch ex As Exception
            Exit Sub
        End Try


        ' Exit , wenn es sich um eine Dependency handelt 
        If shapeType = PTshty.dependency Or shapeType = PTshty.status Then
            Exit Sub
        End If

        ' damit curCoord korrekt bestimmt werden kann, muss im Falle Phase / Meilenstein im Extended Mode
        ' die Gruppierung des Projekt-Shapes erst aufgehoben werden 
        If Not isProjectType(shapeType) Then

            shapeSammlung = ungroupShapes(pName)
            notRegroupedAgain = True

        End If

        With shpElement
            curCoord(0) = .Top
            curCoord(1) = .Left
            curCoord(2) = .Height
            curCoord(3) = .Width
        End With


        ' Rotation ist überhaupt nicht zugelassen 
        If shpElement.Rotation <> 0 Then
            shpElement.Rotation = 0
        End If


        ' jetzt werden die Prüfungen durchgeführt
        If arraysAreDifferent(curCoord, oldCoord) Then

            ' die Height muss immer gleich bleiben 
            shpElement.Height = CSng(oldCoord(2))


            If moveAllowed Then

                ' top darf nur bei ProjektE oder ProjektN verändert werden 
                If curCoord(0) <> oldCoord(0) Then

                    If isProjectType(shapeType) Then
                        Dim tmpZeile As Integer = calcYCoordToZeile(curCoord(0))
                        shpElement.Top = CSng(calcZeileToYCoord(tmpZeile))
                        curCoord(0) = shpElement.Top
                    Else
                        ' korrigiere die Position 
                        shpElement.Top = CSng(oldCoord(0))
                        curCoord(0) = shpElement.Top
                    End If


                End If


                If isProjectType(shapeType) Then
                    ' für Projekte: berechne das neue Start-Datum und ggf die neue Dauer
                    Dim newStartdate As Date
                    Dim newEndDate As Date
                    Dim tmpDauerIndays = hproj.dauerInDays
                    newProjekt = New clsProjekt


                    ' wenn gedehnt bzw. gestaucht wird ...
                    If curCoord(3) <> oldCoord(3) Then
                        ' es wird gestaucht bzw. gedehnt
                        newStartdate = hproj.startDate.AddDays(calcXCoordToTage(curCoord(1) - oldCoord(1)))
                        newEndDate = newStartdate.AddDays(hproj.dauerInDays - 1 + calcXCoordToTage(curCoord(3) - oldCoord(3)))
                    Else

                        ' es wurde nur verschoben 
                        newStartdate = hproj.startDate.AddDays(calcXCoordToTage(curCoord(1) - oldCoord(1)))
                        newEndDate = newStartdate.AddDays(hproj.dauerInDays - 1)

                        Dim newZeile As Integer = calcYCoordToZeile(shpElement.Top)
                        Dim anzahlZeilen As Integer = getNeededSpace(shpElement)

                        ' Platz schaffen auf der Projekt-Tafel
                        If Not magicBoardIstFrei(mycollection:=selCollection, pname:=hproj.name, zeile:=newZeile, _
                                            spalte:=hproj.Start, laenge:=hproj.anzahlRasterElemente, _
                                            anzahlZeilen:=anzahlZeilen) Then

                            If curCoord(0) < oldCoord(0) Then
                                ' es wurde nach oben verschoben - der unten frei werdende Platz kann gnutzt werden 
                                ' alle darunter ligenden Shapes müssen nicht weiter nach unten verschoben werden 
                                Dim stoppzeile As Integer = calcYCoordToZeile(oldCoord(0))
                                Call moveShapesDown(selCollection, newZeile, anzahlZeilen, stoppzeile)
                            Else
                                Call moveShapesDown(selCollection, newZeile, anzahlZeilen, 0)
                            End If

                        End If

                        ' tfzeile setzen
                        hproj.tfZeile = newZeile


                    End If

                    'hproj.copyAttrTo(newProjekt)
                    hproj.korrCopyTo(newProjekt, newStartdate, newEndDate)
                    With hproj
                        newProjekt.name = .name
                        newProjekt.variantName = .variantName
                        newProjekt.ampelStatus = .ampelStatus
                        newProjekt.ampelErlaeuterung = .ampelErlaeuterung
                        newProjekt.Status = .Status
                        newProjekt.shpUID = .shpUID
                        newProjekt.tfZeile = .tfZeile
                    End With

                    newProjekt.timeStamp = Date.Now
                    ' Workaround: 
                    Dim tmpValue As Integer = newProjekt.dauerInDays
                    Call awinCreateBudgetWerte(newProjekt)

                    ' jetzt muss das Projekt aus der Showprojekte und der AlleProjekte herausgenommen werden 
                    ' und in der kopierten Form wieder aufgenommen werden 
                    Dim key As String = pName
                    ShowProjekte.Remove(pName)
                    key = calcProjektKey(hproj)
                    AlleProjekte.Remove(key)

                    AlleProjekte.Add(key, newProjekt)
                    ShowProjekte.Add(newProjekt)

                    Dim zeile As Integer = calcYCoordToZeile(shpElement.Top)


                    pShape = shpElement
                    Dim phaseList As Collection
                    Dim milestoneList As Collection
                    Dim typCollection As New Collection
                    typCollection.Add(CInt(PTshty.phaseN).ToString, CInt(PTshty.phaseN).ToString)
                    typCollection.Add(CInt(PTshty.phaseE).ToString, CInt(PTshty.phaseE).ToString)
                    phaseList = projectboardShapes.getAllChildswithType(pShape, typCollection)

                    typCollection.Clear()
                    typCollection.Add(CInt(PTshty.milestoneN).ToString, CInt(PTshty.milestoneN).ToString)
                    typCollection.Add(CInt(PTshty.milestoneE).ToString, CInt(PTshty.milestoneE).ToString)
                    milestoneList = projectboardShapes.getAllChildswithType(pShape, typCollection)

                    Call clearProjektinPlantafel(pName)
                    ' in selCollection sind die Namen der Projekte, die beim Neuzeichnen nicht berücksichtigt werden sollen, weil 
                    ' sie noch in der Select Collection sind und danach noch behandelt werden 
                    Dim tmpCollection As New Collection
                    Call ZeichneProjektinPlanTafel(noCollection:=selCollection, pname:=newProjekt.name, tryzeile:=hproj.tfZeile, _
                                                   drawPhaseList:=phaseList, drawMilestoneList:=milestoneList)

                    ' Shape wurde gelöscht , der Variable shpElement muss das neue Shape wieder zugewiesen werden 
                    ' damit die aufrufende Routine das shpelement wieder hat 
                    tmpRange = CType(appInstance.Worksheets(arrWsNames(3)), Excel.Worksheet).Shapes.Range(pName)
                    shpElement = tmpRange.Item(1)

                    ' workaround: 
                    tmpDauerIndays = hproj.dauerInDays
                    Call awinCreateBudgetWerte(hproj)


                ElseIf shapeType = PTshty.phaseE Or shapeType = PTshty.phaseN Then
                    ' für Phasen: berechne das neue Start-Datum und ggf. die neue Dauer (muss innerhalb Projekt bleiben !
                    Dim cphase As clsPhase
                    Dim projectBorderLinks As Double = calcDateToXCoord(hproj.startDate)
                    Dim projectBorderRechts As Double = calcDateToXCoord(hproj.startDate.AddDays(hproj.dauerInDays - 1))
                    Dim offsetinTagen As Integer, dauerinTagen As Integer
                    Dim reDraw As Boolean = False
                    Dim tmpDauerIndays = hproj.dauerInDays
                    Dim diffDays As Integer = 0

                    phaseName = extractName(shpElement.Name, PTshty.phaseN)
                    cphase = hproj.getPhase(phaseName)

                   

                    If cphase.name = hproj.name Then
                        ' hier muss die Sonderbehandlung der Phase 1 rein' sicherstellen, 
                        ' daß die Phase 1 in curCoord die richtigen Koordinaten hat 
                        ' und dass die notwendigen Anpassungen der anderen Phasen gemacht wurde 
                        Dim phBorderLinks As Double = phasesBorderLinks(hproj)
                        Dim phBorderRechts As Double = phasesBorderRechts(hproj)

                        ' ist der linke Rand ok? 
                        If curCoord(1) < phBorderLinks Then
                            If curCoord(1) + curCoord(3) >= phBorderRechts Then
                                ' alles ok
                            Else
                                curCoord(1) = phBorderRechts - curCoord(3)
                                reDraw = True
                            End If
                        Else
                            curCoord(1) = phBorderLinks
                            reDraw = True
                        End If


                        ' ist der Rechte Rand ok? 
                        If curCoord(1) + curCoord(3) >= phBorderRechts Then
                            ' alles ok 
                        Else
                            curCoord(3) = phBorderRechts - curCoord(1)
                            reDraw = True
                        End If

                        ' jetzt enthalten die CurCoord die exakten Daten
                        ' bei Phase 1 ist der Offset immer Null aber die diffdays zur Anpassung der 
                        ' Offsets der anderen Phasen müssen gesetzt werden 
                        diffDays = cphase.startOffsetinDays + calcXCoordToTage(curCoord(1) - oldCoord(1))
                        dauerinTagen = cphase.dauerInDays + calcXCoordToTage(curCoord(3) - oldCoord(3))
                        offsetinTagen = 0
                        If diffDays <> 0 Then
                            hproj.startDate = hproj.startDate.AddDays(diffDays)
                            Call hproj.syncXWertePhases()
                        End If



                    Else
                        ' befindet sich die Shape noch innerhalb der Projekt-Grenzen 
                        If curCoord(1) < projectBorderLinks Then
                            If curCoord(3) <> oldCoord(3) Then
                                ' es wurde gedehnt
                                curCoord(3) = curCoord(3) - (projectBorderLinks - curCoord(1))
                            End If
                            curCoord(1) = projectBorderLinks
                            reDraw = True
                        End If

                        If curCoord(1) > projectBorderRechts Then
                            ' gar nicht zugelassen
                            curCoord(1) = oldCoord(1)
                            reDraw = True
                        End If

                        If curCoord(1) + curCoord(3) > projectBorderRechts Then
                            ' dann muss die Breite angepasst werden 
                            curCoord(3) = projectBorderRechts - curCoord(1)
                            reDraw = True
                        End If

                        ' jetzt enthalten die CurCoord die exakten Daten  
                        offsetinTagen = cphase.startOffsetinDays + calcXCoordToTage(curCoord(1) - oldCoord(1))
                        dauerinTagen = cphase.dauerInDays + calcXCoordToTage(curCoord(3) - oldCoord(3))

                    End If


                    If offsetinTagen <> cphase.startOffsetinDays Or dauerinTagen <> cphase.dauerInDays Or _
                        diffDays <> 0 Then
                        Dim faktor As Double = dauerinTagen / cphase.dauerInDays

                        reDraw = True

                        Call cphase.changeStartandDauer(offsetinTagen, dauerinTagen)
                        If faktor <> 1.0 Then
                            ' es wurde gedehnt oder gestaucht, d.h die Meilensteine müssen entsprechend angepasst werden 
                            Call cphase.adjustMilestones(faktor)
                        End If

                        If cphase.name = hproj.name Then
                            ' in diesem Fall wurde die Phase 1 verändert - wenn sich der line Rand der 
                            ' Phase 1 verändert hat, müssen die Pahsen 2 bis N ihren Startoffsets neu berechnet werden 
                            If curCoord(1) <> oldCoord(1) Then
                                reDraw = True
                                Call reCalcOffsetInPhases(hproj, diffDays)
                            End If
                        End If


                    End If

                    If reDraw Then
                        ' es gab Änderungen , zugelassen oder nicht: deshalb muss das Shape neu gezeichnet werden 
                        Call reGroupShape(shapeSammlung, pName)

                        ' pshape ist das übergeordnete Shpelement 
                        pShape = ShowProjekte.getShape(hproj.name)

                        Dim phaseList As Collection
                        Dim milestoneList As Collection
                        Dim typCollection As New Collection
                        typCollection.Add(CInt(PTshty.phaseN).ToString, CInt(PTshty.phaseN).ToString)
                        typCollection.Add(CInt(PTshty.phaseE).ToString, CInt(PTshty.phaseE).ToString)
                        phaseList = projectboardShapes.getAllChildswithType(pShape, typCollection)

                        typCollection.Clear()
                        typCollection.Add(CInt(PTshty.milestoneN).ToString, CInt(PTshty.milestoneN).ToString)
                        typCollection.Add(CInt(PTshty.milestoneE).ToString, CInt(PTshty.milestoneE).ToString)
                        milestoneList = projectboardShapes.getAllChildswithType(pShape, typCollection)


                        Call clearProjektinPlantafel(pName)
                        ' wenn bestimmte Projekte beim Suchen nach einem Platz nicht berücksichtigt werden sollen,
                        ' dann müssen sie in einer Collection an ZeichneProjektinPlanTafel übergeben werden 
                        Dim tmpCollection As New Collection
                        Call ZeichneProjektinPlanTafel(noCollection:=tmpCollection, pname:=pName, tryzeile:=hproj.tfZeile, _
                                                       drawPhaseList:=phaseList, drawMilestoneList:=milestoneList)
                        notRegroupedAgain = False

                        ' Shape-Element wurde gelöscht , jetzt muss dem shpElement wieder das entsprechende 
                        ' Projekt-Shape zugewiesen werden 
                        tmpRange = CType(appInstance.Worksheets(arrWsNames(3)), Excel.Worksheet).Shapes.Range(pName)
                        shpElement = tmpRange.Item(1)

                        ' jetzt noch die Budget Werte neu berechnen 
                        ' Workaround: 
                        Dim tmpValue As Integer = hproj.dauerInDays
                        Call awinCreateBudgetWerte(hproj)

                    End If


                ElseIf shapeType = PTshty.milestoneN Or shapeType = PTshty.milestoneE Then
                    ' für Meilensteine: berechne das neue Datum ; muss innerhalb der Phase bleiben 

                    Dim cphase As clsPhase
                    Dim cresult As clsMeilenstein

                    Dim reDraw As Boolean = False
                    Dim tmpDauerIndays = hproj.dauerInDays
                    Dim diffDays As Integer = 0

                    reDraw = False
                    phaseName = extractName(shpElement.Name, PTshty.phaseN)
                    resultNr = CInt(extractName(shpElement.Name, PTshty.milestoneN))
                    cphase = hproj.getPhase(phaseName)
                    cresult = cphase.getResult(resultNr)

                    Dim phBorderLinks As Double = calcDateToXCoord(cphase.getStartDate)
                    Dim phBorderRechts As Double = calcDateToXCoord(cphase.getEndDate)

                    diffDays = cphase.startOffsetinDays + calcXCoordToTage(curCoord(1) - oldCoord(1))

                    ' ist der linke Rand ok ? 
                    If curCoord(1) + curCoord(3) / 2 < phBorderLinks Then
                        ' out of bound
                        curCoord(1) = phBorderLinks - curCoord(3) / 2
                        reDraw = True
                    End If

                    ' ist der rechte Rand ok ? 
                    If curCoord(1) + curCoord(3) / 2 > phBorderRechts Then
                        ' out of bound
                        curCoord(1) = phBorderRechts - curCoord(3) / 2
                        reDraw = True
                    End If

                    ' jetzt ist sichergestellt, daß eine gültige Position gefunden ist 
                    Dim newDate As Date = calcXCoordToDate(curCoord(1) + curCoord(3) / 2)
                    If DateDiff(DateInterval.Day, newDate, cresult.getDate) <> 0 Then
                        cresult.setDate = newDate
                        reDraw = True
                    End If

                    If reDraw Then
                        ' es gab Änderungen , zugelassen oder nicht: deshalb muss das Shape neu gezeichnet werden 
                        Call reGroupShape(shapeSammlung, pName)

                        ' pshape ist das übergeordnete Shpelement 
                        pShape = ShowProjekte.getShape(hproj.name)

                        Dim phaseList As Collection
                        Dim milestoneList As Collection
                        Dim typCollection As New Collection
                        typCollection.Add(CInt(PTshty.phaseN).ToString, CInt(PTshty.phaseN).ToString)
                        typCollection.Add(CInt(PTshty.phaseE).ToString, CInt(PTshty.phaseE).ToString)
                        phaseList = projectboardShapes.getAllChildswithType(pShape, typCollection)

                        typCollection.Clear()
                        typCollection.Add(CInt(PTshty.milestoneN).ToString, CInt(PTshty.milestoneN).ToString)
                        typCollection.Add(CInt(PTshty.milestoneE).ToString, CInt(PTshty.milestoneE).ToString)
                        milestoneList = projectboardShapes.getAllChildswithType(pShape, typCollection)

                        Call clearProjektinPlantafel(pName)

                        ' wenn bestimmte Projekte beim Suchen nach einem Platz nicht berücksichtigt werden sollen,
                        ' dann müssen sie in einer Collection an ZeichneProjektinPlanTafel übergeben werden 
                        Dim tmpCollection As New Collection
                        Call ZeichneProjektinPlanTafel(noCollection:=tmpCollection, pname:=pName, tryzeile:=hproj.tfZeile, _
                                                       drawPhaseList:=phaseList, drawMilestoneList:=milestoneList)
                        notRegroupedAgain = False

                        ' Shape-Element wurde gelöscht , jetzt muss dem shpElement wieder das entsprechende 
                        ' Projekt-Shape zugewiesen werden 
                        tmpRange = CType(appInstance.Worksheets(arrWsNames(3)), Excel.Worksheet).Shapes.Range(pName)
                        shpElement = tmpRange.Item(1)

                    End If


                End If


            Else
                ' auf die alten Koordinaten zurücksetzen 

                ' es gab Änderungen , zugelassen oder nicht: deshalb muss das Shape neu gezeichnet werden 
                Dim newZeile As Integer = hproj.tfZeile

                ' den Befel braucht man, damit später alle Shapes auf einen Schlag gelöscht 
                ' und neu gezeichnet werden können 
                If notRegroupedAgain And Not isProjectType(shapeType) Then
                    Call reGroupShape(shapeSammlung, pName)
                End If


                ' top darf nur bei ProjektE, ProjektC oder ProjektN verändert werden 
                If curCoord(0) <> oldCoord(0) Then

                    If isProjectType(shapeType) Then
                        newZeile = calcYCoordToZeile(curCoord(0))

                        ' Platz schaffen auf der Projekt-Tafel

                        Dim anzahlZeilen As Integer = getNeededSpace(hproj)
                        If Not magicBoardIstFrei(mycollection:=selCollection, pname:=hproj.name, zeile:=newZeile, _
                                            spalte:=hproj.Start, laenge:=hproj.anzahlRasterElemente, _
                                            anzahlZeilen:=anzahlZeilen) Then

                            If curCoord(0) < oldCoord(0) Then
                                ' es wurde nach oben verschoben - der unten frei werdende Platz kann gnutzt werden 
                                ' alle darunter ligenden Shapes müssen nicht weiter nach unten verschoben werden 
                                Dim stoppzeile As Integer = calcYCoordToZeile(oldCoord(0))
                                Call moveShapesDown(selCollection, newZeile, anzahlZeilen, stoppzeile)
                            Else
                                Call moveShapesDown(selCollection, newZeile, anzahlZeilen, 0)
                            End If


                        End If


                    Else
                        ' korrigiere die Höhe 
                        newZeile = hproj.tfZeile
                    End If


                End If

                ' jetzt muss ggf das übergeordnete Projektshape geholt werden 
                If isProjectType(shapeType) Then
                    pShape = shpElement
                Else
                    pShape = ShowProjekte.getShape(hproj.name)
                End If

                Dim typCollection As New Collection
                typCollection.Add(CInt(PTshty.phaseN).ToString, CInt(PTshty.phaseN).ToString)
                typCollection.Add(CInt(PTshty.phaseE).ToString, CInt(PTshty.phaseE).ToString)
                Dim phaseList As Collection = Me.getAllChildswithType(pShape, typCollection)

                typCollection.Clear()
                typCollection.Add(CInt(PTshty.milestoneN).ToString, CInt(PTshty.milestoneN).ToString)
                typCollection.Add(CInt(PTshty.milestoneE).ToString, CInt(PTshty.milestoneE).ToString)
                Dim milestoneList As Collection = Me.getAllChildswithType(pShape, typCollection)


                Call clearProjektinPlantafel(pName)
                ' in selCollection sind die Namen der Projekte, die beim Neuzeichnen nicht berücksichtigt werden sollen, weil 
                ' sie noch in der Select Collection sind und danach noch behandelt werden  
                Call ZeichneProjektinPlanTafel(noCollection:=selCollection, pname:=pName, tryzeile:=newZeile, _
                                                drawPhaseList:=phaseList, drawMilestoneList:=milestoneList)
                notRegroupedAgain = False

                ' Shape-Element wurde gelöscht , jetzt muss dem shpElement wieder das entsprechende 
                ' Projekt-Shape zugewiesen werden 
                tmpRange = CType(appInstance.Worksheets(arrWsNames(3)), Excel.Worksheet).Shapes.Range(pName)
                shpElement = tmpRange.Item(1)

            End If

        End If

        ' jetzt wird wieder regruppiert
        If (shapeType = PTshty.phaseE Or shapeType = PTshty.milestoneE) And notRegroupedAgain Then

            Call reGroupShape(shapeSammlung, pName)

        End If

    End Sub

    Public Sub New()
        AllShapes = New SortedList(Of String, Double())
    End Sub

    '  
    ''' <summary>
    ''' gibt den Wert in X-Koordinaten aus, der den linken Rand von Phase 2 bis N bedeutet
    ''' -1 wenn es keine Phasen gibt  
    ''' </summary>
    ''' <param name="project"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Function phasesBorderLinks(ByVal project As clsProjekt) As Double
        Dim tmpValue As Double = -1.0
        Dim curValue As Double

        For p = 2 To project.CountPhases
            curValue = calcDateToXCoord(project.getPhase(p).getStartDate)
            If p = 2 Then
                tmpValue = curValue
            ElseIf curValue < tmpValue Then
                tmpValue = curValue
            End If
        Next

        phasesBorderLinks = tmpValue

    End Function

    ''' <summary>
    ''' gibt den Wert in X-Koordinaten aus, der den rechten Rand von Phase 2 bis N bedeutet
    ''' -1 wenn es keine Phasen gibt 
    ''' </summary>
    ''' <param name="project"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Function phasesBorderRechts(ByVal project As clsProjekt) As Double
        Dim tmpValue As Double = -1.0
        Dim curValue As Double

        For p = 2 To project.CountPhases
            curValue = calcDateToXCoord(project.getPhase(p).getEndDate)
            If p = 2 Then
                tmpValue = curValue
            ElseIf curValue > tmpValue Then
                tmpValue = curValue
            End If
        Next

        phasesBorderRechts = tmpValue

    End Function

    ''' <summary>
    ''' der Phasenoffset wird neu berechnet, da sich die Phase 1 als Repräsentant des Projekts verschoben hat 
    ''' </summary>
    ''' <param name="diffDays"></param>
    ''' <remarks></remarks>
    Private Sub reCalcOffsetInPhases(ByRef project As clsProjekt, ByVal diffDays As Integer)
        Dim newOffset As Integer, oldDauer As Integer
        Dim cphase As clsPhase

        For p = 2 To project.CountPhases

            cphase = project.getPhase(p)
            newOffset = cphase.startOffsetinDays
            oldDauer = cphase.dauerInDays

            If newOffset - diffDays < 0 Then
                newOffset = 0
            Else
                newOffset = newOffset - diffDays
            End If

            Call cphase.changeStartandDauer(newOffset, oldDauer)

        Next

    End Sub

End Class
