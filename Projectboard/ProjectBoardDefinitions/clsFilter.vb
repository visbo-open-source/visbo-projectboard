Public Class clsFilter


    Private filterPhase As Collection
    Private filterMilestone As Collection
    Private filterRolle As Collection
    Private filterCost As Collection
    Private filterTyp As Collection
    Private filterBU As Collection
    Private _name As String


    ''' <summary>
    ''' prüft ob irgendein Filter gesetzt ist 
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public ReadOnly Property isEmpty As Boolean
        Get
            Dim sum As Integer = filterPhase.Count + filterMilestone.Count + _
                                 filterRolle.Count + filterCost.Count + _
                                 filterTyp.Count + filterBU.Count
            If sum = 0 Then
                isEmpty = True
            Else
                isEmpty = False
            End If

        End Get
    End Property

    ''' <summary>
    ''' schreibt/liest die Filter Collection der BUs
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Property BUs() As Collection
        Get
            BUs = filterBU
        End Get
        Set(value As Collection)

            If Not IsNothing(value) Then
                filterBU = value
            End If

        End Set
    End Property

    ''' <summary>
    ''' schreibt/liest die Filter Collection der Typen
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Property Typs() As Collection
        Get
            Typs = filterTyp
        End Get
        Set(value As Collection)

            If Not IsNothing(value) Then
                filterTyp = value
            End If

        End Set
    End Property

    ''' <summary>
    ''' schreibt/liest die Filter Collection der Phasen
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Property Phases() As Collection
        Get
            Phases = filterPhase
        End Get
        Set(value As Collection)

            If Not IsNothing(value) Then
                filterPhase = value
            End If

        End Set
    End Property

    ''' <summary>
    ''' schreibt/liest die Filter Collection der Meilensteine
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Property Milestones() As Collection
        Get
            Milestones = filterMilestone
        End Get
        Set(value As Collection)

            If Not IsNothing(value) Then
                filterMilestone = value
            End If

        End Set
    End Property

    ''' <summary>
    ''' schreibt/liest die Filter Collection der Rolle
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Property Roles() As Collection
        Get
            Roles = filterRolle
        End Get
        Set(value As Collection)

            If Not IsNothing(value) Then
                filterRolle = value
            End If

        End Set
    End Property

    ''' <summary>
    ''' schreibt/liest die Filter Collection der Kostenart
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Property Costs() As Collection
        Get
            Costs = filterCost
        End Get
        Set(value As Collection)

            If Not IsNothing(value) Then
                filterCost = value
            End If

        End Set
    End Property


    ''' <summary>
    ''' liest bzw. schreibt den Namen des filters 
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Property name As String

        Get
            name = _name
        End Get
        Set(value As String)

            If Not IsNothing(value) Then
                If value.Trim.Length > 0 Then
                    _name = value
                Else
                    _name = "XXX"
                End If
            Else
                _name = "XXX"
            End If

        End Set
    End Property

    

    ''' <summary>
    ''' fügt dem Business Unit Filter einen Eintrag hinzu
    ''' wenn businessUnit bereits im Filter vorhanden ist, dann wird nichts hinzugefügt
    ''' aber auch keine Fehlermeldung geworfen 
    ''' </summary>
    ''' <param name="businessUnit"></param>
    ''' <remarks></remarks>
    Public Sub addBU(ByVal businessUnit As String)

        If filterBU.Contains(businessUnit) Then
            ' nichts tun ..
        Else

            If Not IsNothing(businessUnit) Then
                filterBU.Add(businessUnit, businessUnit)
            End If

        End If

    End Sub

    ''' <summary>
    ''' entfernt aus dem Business Unit Filter einen Eintrag
    ''' wenn der Eintrag nicht vorhanden ist, wird nichts entfernt
    ''' aber auch keine Fehlermeldung geworfen 
    ''' </summary>
    ''' <param name="businessUnit"></param>
    ''' <remarks></remarks>
    Public Sub removeBU(ByVal businessUnit As String)

        If Not IsNothing(businessUnit) Then
            If filterBU.Contains(businessUnit) Then
                filterBU.Remove(businessUnit)
            Else
                ' nichts tun ..
            End If
        End If

    End Sub

    ''' <summary>
    ''' gibt true zurück , wenn das Projekt 
    ''' 1. zu einer der im filterBU angegebenen BUs gehört,  und
    ''' 2. zu einem der im filterTyp angegebenen Projekttypen gehört,  und 
    ''' 3. wenigstens einen der angegebenen Meilensteine - im angegebenen Zeitraum - enthält , oder
    ''' 4. wenigstens eine der angegebenen Phasen - im angegebenen Zeitraum - enthält 
    ''' </summary>
    ''' <param name="hproj"></param>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public ReadOnly Property doesNotBlock(ByVal hproj As clsProjekt) As Boolean
        Get
            Dim containsBU As Boolean
            Dim containsTyp As Boolean
            Dim containsMS As Boolean
            Dim containsPH As Boolean
            Dim containsRole As Boolean
            Dim containsCost As Boolean
            Dim stillOK As Boolean
            Dim tmpMilestone As clsMeilenstein
            Dim tmpPhase As clsPhase
            Dim ix As Integer
            Dim fullName As String

            If Not IsNothing(Me) Then

                ' Überprüfe BU 
                If filterBU.Count = 0 Then
                    containsBU = True

                ElseIf filterBU.Count > 0 And Not IsNothing(hproj.businessUnit) Then

                    If hproj.businessUnit.Trim.Length > 0 Then
                        If filterBU.Contains(hproj.businessUnit.Trim) Then
                            containsBU = True
                        Else
                            containsBU = False
                        End If
                    Else
                        If filterBU.Contains("unknown") Then
                            containsBU = True
                        Else
                            containsBU = False
                        End If
                    End If

                ElseIf IsNothing(hproj.businessUnit) Then
                    If filterBU.Contains("unknown") Then
                        containsBU = True
                    Else
                        containsBU = False
                    End If

                End If


                stillOK = containsBU

                ' überprüfe Typ

                If stillOK Then
                    If filterTyp.Count = 0 Then
                        containsTyp = True

                    ElseIf filterTyp.Count > 0 And Not IsNothing(hproj.VorlagenName) Then
                        If hproj.VorlagenName.Trim.Length > 0 Then
                            If filterTyp.Contains(hproj.VorlagenName.Trim) Then
                                containsTyp = True
                            Else
                                containsTyp = False
                            End If
                        Else
                            If filterTyp.Contains("unknown") Then
                                containsTyp = True
                            Else
                                containsTyp = False
                            End If
                        End If

                    ElseIf IsNothing(hproj.VorlagenName) Then
                        If filterTyp.Contains("unknown") Then
                            containsTyp = True
                        Else
                            containsTyp = False
                        End If

                    Else
                        containsTyp = True

                    End If
                    stillOK = containsTyp
                End If

                If stillOK Then
                    ' Überprüfen Meilensteine und Phasen
                    ' Änderung tk: wenn ein Projekt nur die Phase(1) enthält und keinerlei Meilensteine, dann kann es nicht angezeigt werden, 
                    ' da bisher nur über die positive Schnittmenge entschieden wird; 
                    ' ggf muss das noch über eine separate Variable entschieden werden ... das könnte eigentlich über die Projektlinie gemacht werden 

                    If hproj.hierarchy.count = 1 And awinSettings.mppProjectsWithNoMPmayPass Then
                        ' nur dann kann es sich ggf um ein leeres Projekt handeln 
                        containsMS = True

                    ElseIf filterMilestone.Count = 0 Then

                        If filterPhase.Count = 0 Then
                            containsMS = True
                        Else
                            containsMS = False
                        End If

                    Else
                        containsMS = False
                        ix = 1

                        While ix <= filterMilestone.Count And Not containsMS

                            fullName = CStr(filterMilestone.Item(ix))
                            Dim curMsName As String = ""
                            Dim breadcrumb As String = ""

                            ' hier wird der Eintrag in filterMilestone aufgesplittet in curMsName und breadcrumb) 
                            Call splitHryFullnameTo2(fullName, curMsName, breadcrumb)

                            Dim milestoneIndices(,) As Integer = hproj.hierarchy.getMilestoneIndices(curMsName, breadcrumb)
                            ' in milestoneIndices sind jetzt die Phasen- und Meilenstein Index der Phasen bzw Meilenstein Liste

                            For mx As Integer = 0 To CInt(milestoneIndices.Length / 2) - 1

                                tmpMilestone = hproj.getMilestone(milestoneIndices(0, mx), milestoneIndices(1, mx))
                                If IsNothing(tmpMilestone) Then



                                Else

                                    If showRangeLeft > 0 And showRangeRight > 0 And showRangeRight >= showRangeLeft Then
                                        ' jetzt muss geprüft werden, ob der Meilenstein auch im angegebenen Bereich liegt 
                                        Dim tmpMsDate As Integer = getColumnOfDate(tmpMilestone.getDate)
                                        If tmpMsDate >= showRangeLeft And tmpMsDate <= showRangeRight Then
                                            containsMS = True
                                            Exit For
                                        End If
                                    Else
                                        containsMS = True
                                        Exit For
                                    End If

                                End If


                            Next

                            ix = ix + 1

                        End While

                    End If

                    ' jetzt werden die Phasen überprüft, aber nur , wenn nicht containsMS bereits true ist 
                    containsPH = False

                    If Not containsMS Then
                        ' prüfe Phasen ; das wird mit Not Stillok geprüft, da es um Meilensteine oder Phasen geht 
                        ' wenn es bereits einen der Meilensteine enthält, ist nicht mehr auf Phasen zu prüfen 
                        If filterPhase.Count = 0 Then
                            containsPH = False
                        Else
                            ix = 1

                            While ix <= filterPhase.Count And Not containsPH

                                fullName = CStr(filterPhase.Item(ix))
                                Dim pName As String = ""
                                Dim breadcrumb As String = ""

                                ' hier wird der Eintrag in filterMilestone aufgesplittet in curMsName und breadcrumb) 
                                Call splitHryFullnameTo2(fullName, pName, breadcrumb)

                                Dim phaseIndices() As Integer = hproj.hierarchy.getPhaseIndices(pName, breadcrumb)

                                For px As Integer = 0 To phaseIndices.Length - 1

                                    tmpPhase = hproj.getPhase(phaseIndices(px))

                                    If IsNothing(tmpPhase) Then



                                    Else

                                        If showRangeLeft > 0 And showRangeRight > 0 Then

                                            Dim leftDate As Date = StartofCalendar.AddMonths(showRangeLeft - 1)
                                            Dim rightdate As Date = StartofCalendar.AddMonths(showRangeRight).AddDays(-1)
                                            Dim tmpPhStart As Date = tmpPhase.getStartDate
                                            Dim tmpPhEnde As Date = tmpPhase.getEndDate

                                            If DateDiff(DateInterval.Day, tmpPhEnde, leftDate) > 0 Or _
                                                DateDiff(DateInterval.Day, tmpPhStart, rightdate) < 0 Then

                                            Else
                                                containsPH = True
                                            End If
                                            ' jetzt muss geprüft werden, ob der Meilenstein auch im angegebenen Bereich liegt 

                                        Else
                                            containsPH = True
                                        End If

                                    End If

                                    If containsPH Then
                                        Exit For
                                    End If

                                Next

                                ix = ix + 1

                            End While

                        End If
                    End If


                    stillOK = containsMS Or containsPH

                End If

            Else
                ' wenn der Filter = Nothing
                stillOK = True
            End If

            ' Prüfen ob bestimmte Rollen vorkommen 
            If stillOK Then

                If filterRolle.Count > 0 Then

                    Dim roleName As String
                    Dim rollenBedarfe As Double = 0.0
                    Dim myCollection As New Collection
                    ' DiagrammTypen(1) = Rollen 
                    Dim type As String = DiagrammTypen(1)
                    ix = 1
                    containsRole = False

                    While ix <= filterRolle.Count And Not containsRole

                        roleName = CStr(filterRolle.Item(ix))

                        ' zurücksetzen
                        myCollection.Clear()
                        rollenBedarfe = 0.0

                        ' berechnen
                        myCollection.Add(roleName, roleName)
                        rollenBedarfe = hproj.getBedarfeInMonths(myCollection, type).Sum

                        ' entscheiden
                        If rollenBedarfe > 0 Then
                            containsRole = True
                        Else
                            ix = ix + 1
                        End If


                    End While

                Else
                    containsRole = True
                End If
                stillOK = containsRole
            End If

            ' Prüfen ob bestimmte Kostenarten vorkommen 
            If stillOK Then

                If filterCost.Count > 0 Then

                    Dim costName As String
                    Dim costBedarfe As Double = 0.0
                    Dim myCollection As New Collection
                    ' DiagrammTypen(1) = Rollen 
                    Dim type As String = DiagrammTypen(2)
                    ix = 1
                    containsCost = False

                    While ix <= filterCost.Count And Not containsCost

                        costName = CStr(filterCost.Item(ix))

                        ' zurücksetzen
                        myCollection.Clear()
                        costBedarfe = 0.0

                        ' berechnen
                        myCollection.Add(costName, costName)
                        costBedarfe = hproj.getBedarfeInMonths(myCollection, type).Sum

                        ' entscheiden
                        If costBedarfe > 0 Then
                            containsCost = True
                        Else
                            ix = ix + 1
                        End If


                    End While

                Else
                    containsCost = True
                End If
                stillOK = containsCost
            End If


            doesNotBlock = stillOK

        End Get
    End Property

    

    Sub New()
        filterBU = New Collection
        filterPhase = New Collection
        filterMilestone = New Collection
        filterTyp = New Collection
        filterRolle = New Collection
        filterCost = New Collection
        _name = "XXX"
    End Sub

    ''' <summary>
    ''' legt einen neuen filter an unter Angabe der bekannten Filter Collections
    ''' Eingabe Parameter kann auch Nothing sein 
    ''' </summary>
    ''' <param name="kennung">Name des Filters</param>
    ''' <param name="fBU">filter BU</param>
    ''' <param name="fTyp">filter Typ</param>
    ''' <param name="fPhase">filter Phase</param>
    ''' <param name="fMilestone">filter Meilenstein</param>
    ''' <param name="fRolle">filter Rolle</param>
    ''' <param name="fCost">filter Cost</param>
    ''' <remarks></remarks>
    Sub New(ByVal kennung As String, _
                ByVal fBU As Collection, ByVal fTyp As Collection, _
                ByVal fPhase As Collection, ByVal fMilestone As Collection, _
                ByVal fRolle As Collection, ByVal fCost As Collection)

        filterPhase = New Collection
        filterPhase = copyCollection(fPhase)

        filterMilestone = New Collection
        filterMilestone = copyCollection(fMilestone)

        filterRolle = New Collection
        filterRolle = copyCollection(fRolle)
        
        filterCost = New Collection
        filterCost = copyCollection(fCost)

        filterBU = New Collection
        filterBU = copyCollection(fBU)
        
        filterTyp = New Collection
        filterTyp = copyCollection(fTyp)
        

        name = kennung

    End Sub
End Class
