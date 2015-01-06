Public Class clsFilter

    Private filterBU As SortedList(Of String, String)
    Private filterPhase As SortedList(Of String, String)
    Private filterMilestone As SortedList(Of String, String)
    Private filterTyp As SortedList(Of String, String)
    Private _name As String
    Private _isActive As Boolean

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
            If value.Trim.Length > 0 Then
                _name = value
            Else
                _name = "XXX"
            End If

        End Set
    End Property

    Public Property isActive As Boolean
        Get
            isActive = _isActive
        End Get
        Set(value As Boolean)
            _isActive = value
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

        If filterBU.ContainsKey(businessUnit) Then
            ' nichts tun ..
        Else
            filterBU.Add(businessUnit, businessUnit)
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

        If filterBU.ContainsKey(businessUnit) Then
            filterBU.Remove(businessUnit)
        Else
            ' nichts tun ..
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
            Dim stillOK As Boolean
            Dim tmpMilestone As clsMeilenstein
            Dim tmpPhase As clsPhase
            Dim ix As Integer


            If _isActive Then

                ' Überprüfe BU 
                If filterBU.Count > 0 Then
                    If hproj.businessUnit.Trim.Length > 0 Then
                        If filterBU.ContainsKey(hproj.businessUnit.Trim) Then
                            containsBU = True
                        Else
                            containsBU = False
                        End If
                    Else
                        containsBU = False
                    End If
                Else
                    containsBU = True
                End If

                stillOK = containsBU

                If stillOK Then
                    If filterTyp.Count > 0 Then
                        If hproj.VorlagenName.Trim.Length > 0 Then
                            If filterTyp.ContainsKey(hproj.VorlagenName.Trim) Then
                                containsTyp = True
                            Else
                                containsTyp = False
                            End If
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
                    If filterMilestone.Count = 0 Then

                        If filterPhase.Count = 0 Then
                            containsMS = True
                        Else
                            containsMS = False
                        End If

                    Else
                        containsMS = False
                        ix = 1

                        While ix <= filterMilestone.Count And Not containsMS
                            tmpMilestone = hproj.getMilestone(filterMilestone.ElementAt(ix - 1).Key)

                            If IsNothing(tmpMilestone) Then

                                ix = ix + 1

                            Else

                                If showRangeLeft > 0 And showRangeRight > 0 Then
                                    ' jetzt muss geprüft werden, ob der Meilenstein auch im angegebenen Bereich liegt 
                                    Dim tmpMsDate As Integer = getColumnOfDate(tmpMilestone.getDate)
                                    If tmpMsDate >= showRangeLeft And tmpMsDate <= showRangeRight Then
                                        containsMS = True
                                    Else
                                        ix = ix + 1
                                    End If
                                Else
                                    containsMS = True
                                End If

                            End If

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
                                tmpPhase = hproj.getPhase(filterPhase.ElementAt(ix - 1).Key)

                                If IsNothing(tmpPhase) Then

                                    ix = ix + 1

                                Else

                                    If showRangeLeft > 0 And showRangeRight > 0 Then

                                        Dim leftDate As Date = StartofCalendar.AddMonths(showRangeLeft - 1)
                                        Dim rightdate As Date = StartofCalendar.AddMonths(showRangeRight).AddDays(-1)
                                        Dim tmpPhStart As Date = tmpPhase.getStartDate
                                        Dim tmpPhEnde As Date = tmpPhase.getEndDate

                                        If DateDiff(DateInterval.Day, tmpPhEnde, leftDate) > 0 Or _
                                            DateDiff(DateInterval.Day, tmpPhStart, rightdate) < 0 Then
                                            containsPH = False
                                        Else
                                            containsPH = True
                                        End If
                                        ' jetzt muss geprüft werden, ob der Meilenstein auch im angegebenen Bereich liegt 

                                    Else
                                        containsPH = True
                                    End If

                                End If

                            End While

                        End If
                    End If


                    stillOK = containsMS Or containsPH

                End If

            Else
                stillOK = True
            End If


            doesNotBlock = stillOK

        End Get
    End Property

    Sub New()
        filterBU = New SortedList(Of String, String)
        filterPhase = New SortedList(Of String, String)
        filterMilestone = New SortedList(Of String, String)
        filterTyp = New SortedList(Of String, String)
    End Sub
End Class
